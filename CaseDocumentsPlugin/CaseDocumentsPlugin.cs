using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CaseDocumentsPlugin
{
    public class CaseDocumentsPlugin : IPlugin
    {
        private static readonly string[] ContactRoles = { "reader", "writer" };

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(null);

            try
            {
                if (string.Equals(context.PrimaryEntityName, "incident", StringComparison.OrdinalIgnoreCase)
                    && context.PrimaryEntityId != Guid.Empty)
                {
                    if (context.MessageName == "Create")
                    {
                        var graph = GraphSession.Create(service);
                        var containerId = EnsureContainer(service, graph, context.PrimaryEntityId);
                        SyncDynamicsAccess(service, tracing, graph, context.PrimaryEntityId, containerId);
                        return;
                    }

                    if (context.MessageName == "Update")
                    {
                        var containerId = StoredContainerId(service, context.PrimaryEntityId);
                        if (!string.IsNullOrWhiteSpace(containerId))
                        {
                            SyncDynamicsAccess(service, tracing, GraphSession.Create(service), context.PrimaryEntityId, containerId);
                        }
                        return;
                    }
                }

                var caseId = RequiredGuid(context, "CaseId");
                var caller = context.InitiatingUserId;
                switch (context.MessageName)
                {
                    case "new_EnsureCaseContainer":
                        context.OutputParameters["Result"] = Ensure(service, caller, caseId);
                        break;
                    case "new_ListCaseAccess":
                        RequireAccess(service, caller, caseId, AccessRights.ReadAccess);
                        context.OutputParameters["Result"] = ListAccess(service, caseId);
                        break;
                    case "new_GrantCaseContactAccess":
                        RequireAccess(service, caller, caseId, AccessRights.WriteAccess);
                        context.OutputParameters["Result"] = GrantContact(
                            service,
                            caseId,
                            RequiredGuid(context, "ContactId"),
                            RequiredRole(context));
                        break;
                    case "new_UpdateCaseAccess":
                        RequireAccess(service, caller, caseId, AccessRights.WriteAccess);
                        context.OutputParameters["Result"] = UpdateContact(
                            service,
                            caseId,
                            Required(context, "PermissionId"),
                            RequiredRole(context));
                        break;
                    case "new_RemoveCaseAccess":
                        RequireAccess(service, caller, caseId, AccessRights.WriteAccess);
                        context.OutputParameters["Result"] = RemoveContact(service, caseId, Required(context, "PermissionId"));
                        break;
                    default:
                        throw new InvalidPluginExecutionException("Unsupported message " + context.MessageName);
                }
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                tracing.Trace(ex.ToString());
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private static string Ensure(IOrganizationService service, Guid caller, Guid caseId)
        {
            var rights = AccessFor(service, caller, caseId);
            if ((rights & AccessRights.ReadAccess) == 0)
            {
                throw new InvalidPluginExecutionException("You do not have access to this case.");
            }

            var graph = GraphSession.Create(service);
            var containerId = EnsureContainer(service, graph, caseId);
            var role = (rights & AccessRights.WriteAccess) != 0 ? "writer" : "reader";
            var upn = UserPrincipalName(service, caller);
            if (!string.IsNullOrWhiteSpace(upn))
            {
                SetMemberRole(graph, containerId, graph.ListPermissions(containerId), upn, role);
            }
            return "{\"containerId\":\"" + JsonText.Escape(containerId) + "\",\"role\":\"" + role + "\"}";
        }

        private static string EnsureContainer(IOrganizationService service, GraphSession graph, Guid caseId)
        {
            var existing = StoredContainerId(service, caseId);
            if (!string.IsNullOrWhiteSpace(existing))
            {
                return existing;
            }

            var incident = service.Retrieve("incident", caseId, new ColumnSet("title", "ticketnumber"));
            var displayName = (incident.GetAttributeValue<string>("ticketnumber") + " " + incident.GetAttributeValue<string>("title")).Trim();
            if (displayName.Length == 0)
            {
                displayName = "Case " + caseId.ToString("N");
            }
            if (displayName.Length > 120)
            {
                displayName = displayName.Substring(0, 120);
            }

            var created = graph.CreateContainer(displayName, caseId.ToString());
            var raced = StoredContainerId(service, caseId);
            if (!string.IsNullOrWhiteSpace(raced))
            {
                graph.DeleteContainer(created);
                return raced;
            }

            var update = new Entity("incident", caseId);
            update["new_sharepointcontainer"] = created;
            service.Update(update);
            return created;
        }

        private static string StoredContainerId(IOrganizationService service, Guid caseId)
        {
            return service.Retrieve("incident", caseId, new ColumnSet("new_sharepointcontainer"))
                .GetAttributeValue<string>("new_sharepointcontainer");
        }

        private static void SyncDynamicsAccess(IOrganizationService service, ITracingService tracing, GraphSession graph, Guid caseId, string containerId)
        {
            var permissions = graph.ListPermissions(containerId);
            foreach (var permission in permissions.Where(p => !p.IsExternal && !string.IsNullOrWhiteSpace(p.UserPrincipalName)))
            {
                var userId = SystemUserId(service, permission.UserPrincipalName);
                var rights = userId == Guid.Empty ? AccessRights.None : AccessFor(service, userId, caseId);
                if ((rights & AccessRights.ReadAccess) == 0)
                {
                    tracing.Trace("Removing {0}: no access to the case in Dynamics.", permission.UserPrincipalName);
                    graph.DeletePermission(containerId, permission.Id);
                    continue;
                }
                var role = (rights & AccessRights.WriteAccess) != 0 ? "writer" : "reader";
                if (!permission.Roles.Contains(role))
                {
                    graph.UpdatePermission(containerId, permission.Id, role);
                }
            }

            permissions = graph.ListPermissions(containerId);
            foreach (var ownerId in CaseOwners(service, caseId))
            {
                var upn = UserPrincipalName(service, ownerId);
                if (!string.IsNullOrWhiteSpace(upn))
                {
                    var rights = AccessFor(service, ownerId, caseId);
                    var role = (rights & AccessRights.WriteAccess) != 0 ? "writer" : "reader";
                    SetMemberRole(graph, containerId, permissions, upn, role);
                }
            }
        }

        private static void SetMemberRole(GraphSession graph, string containerId, List<ContainerPermission> permissions, string upn, string role)
        {
            var current = permissions.FirstOrDefault(p => string.Equals(p.UserPrincipalName, upn, StringComparison.OrdinalIgnoreCase));
            if (current == null)
            {
                graph.GrantPermission(containerId, upn, role);
            }
            else if (!current.Roles.Contains(role) && !current.Roles.Contains("owner") && !current.Roles.Contains("manager"))
            {
                graph.UpdatePermission(containerId, current.Id, role);
            }
        }

        private static IEnumerable<Guid> CaseOwners(IOrganizationService service, Guid caseId)
        {
            var owner = service.Retrieve("incident", caseId, new ColumnSet("ownerid")).GetAttributeValue<EntityReference>("ownerid");
            if (owner == null)
            {
                return Enumerable.Empty<Guid>();
            }
            if (owner.LogicalName == "systemuser")
            {
                return new[] { owner.Id };
            }

            var members = new QueryExpression("teammembership") { ColumnSet = new ColumnSet("systemuserid") };
            members.Criteria.AddCondition("teamid", ConditionOperator.Equal, owner.Id);
            return service.RetrieveMultiple(members).Entities
                .Select(e => e.GetAttributeValue<Guid>("systemuserid"))
                .Where(id => id != Guid.Empty)
                .ToList();
        }

        private static string ListAccess(IOrganizationService service, Guid caseId)
        {
            var containerId = StoredContainerId(service, caseId);
            if (string.IsNullOrWhiteSpace(containerId))
            {
                return "{\"entries\":[]}";
            }
            var entries = GraphSession.Create(service).ListPermissions(containerId);
            return "{\"entries\":[" + string.Join(",", entries.Select(EntryJson)) + "]}";
        }

        private static string GrantContact(IOrganizationService service, Guid caseId, Guid contactId, string role)
        {
            var containerId = StoredContainerId(service, caseId);
            if (string.IsNullOrWhiteSpace(containerId))
            {
                throw new InvalidPluginExecutionException("This case does not have a container yet.");
            }

            var contact = service.Retrieve("contact", contactId, new ColumnSet("fullname", "emailaddress1"));
            var email = (contact.GetAttributeValue<string>("emailaddress1") ?? "").Trim();
            if (email.Length == 0)
            {
                throw new InvalidPluginExecutionException(
                    (contact.GetAttributeValue<string>("fullname") ?? "This contact") + " has no email address.");
            }

            var graph = GraphSession.Create(service);
            var upn = graph.FindOrInviteGuest(email, contact.GetAttributeValue<string>("fullname"));
            var permissions = graph.ListPermissions(containerId);
            var current = permissions.FirstOrDefault(p => string.Equals(p.UserPrincipalName, upn, StringComparison.OrdinalIgnoreCase));
            if (current != null && !current.IsExternal)
            {
                throw new InvalidPluginExecutionException(email + " is a colleague. Their access follows Dynamics.");
            }

            var entry = current == null
                ? graph.GrantPermission(containerId, upn, role)
                : graph.UpdatePermission(containerId, current.Id, role);
            return "{\"entry\":" + EntryJson(entry) + "}";
        }

        private static string UpdateContact(IOrganizationService service, Guid caseId, string permissionId, string role)
        {
            var containerId = StoredContainerId(service, caseId);
            var graph = GraphSession.Create(service);
            RequireExternal(graph, containerId, permissionId);
            return "{\"entry\":" + EntryJson(graph.UpdatePermission(containerId, permissionId, role)) + "}";
        }

        private static string RemoveContact(IOrganizationService service, Guid caseId, string permissionId)
        {
            var containerId = StoredContainerId(service, caseId);
            var graph = GraphSession.Create(service);
            RequireExternal(graph, containerId, permissionId);
            graph.DeletePermission(containerId, permissionId);
            return "{}";
        }

        private static void RequireExternal(GraphSession graph, string containerId, string permissionId)
        {
            if (string.IsNullOrWhiteSpace(containerId))
            {
                throw new InvalidPluginExecutionException("This case does not have a container yet.");
            }
            var permission = graph.ListPermissions(containerId).FirstOrDefault(p => p.Id == permissionId);
            if (permission == null)
            {
                throw new InvalidPluginExecutionException("That access entry no longer exists.");
            }
            if (!permission.IsExternal)
            {
                throw new InvalidPluginExecutionException("Colleague access follows Dynamics and cannot be changed here.");
            }
        }

        private static string EntryJson(ContainerPermission entry)
        {
            return "{\"id\":\"" + JsonText.Escape(entry.Id)
                + "\",\"name\":\"" + JsonText.Escape(entry.DisplayName)
                + "\",\"email\":\"" + JsonText.Escape(entry.Email)
                + "\",\"role\":\"" + JsonText.Escape(entry.Roles.FirstOrDefault() ?? "")
                + "\",\"external\":" + (entry.IsExternal ? "true" : "false") + "}";
        }

        private static void RequireAccess(IOrganizationService service, Guid userId, Guid caseId, AccessRights needed)
        {
            if ((AccessFor(service, userId, caseId) & needed) == 0)
            {
                throw new InvalidPluginExecutionException(
                    needed == AccessRights.WriteAccess
                        ? "You need permission to edit this case to change who has access."
                        : "You do not have access to this case.");
            }
        }

        private static AccessRights AccessFor(IOrganizationService service, Guid userId, Guid caseId)
        {
            var response = (RetrievePrincipalAccessResponse)service.Execute(new RetrievePrincipalAccessRequest
            {
                Principal = new EntityReference("systemuser", userId),
                Target = new EntityReference("incident", caseId)
            });
            return response.AccessRights;
        }

        private static string UserPrincipalName(IOrganizationService service, Guid userId)
        {
            return service.Retrieve("systemuser", userId, new ColumnSet("domainname")).GetAttributeValue<string>("domainname");
        }

        private static Guid SystemUserId(IOrganizationService service, string upn)
        {
            var query = new QueryExpression("systemuser") { ColumnSet = new ColumnSet(false), TopCount = 1 };
            query.Criteria.AddCondition("domainname", ConditionOperator.Equal, upn);
            query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);
            var found = service.RetrieveMultiple(query).Entities;
            return found.Count == 0 ? Guid.Empty : found[0].Id;
        }

        private static string Required(IPluginExecutionContext context, string name)
        {
            var value = context.InputParameters.Contains(name) ? context.InputParameters[name] as string : null;
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidPluginExecutionException(name + " is required.");
            }
            return value.Trim();
        }

        private static Guid RequiredGuid(IPluginExecutionContext context, string name)
        {
            Guid parsed;
            if (!Guid.TryParse(Required(context, name).Trim('{', '}'), out parsed))
            {
                throw new InvalidPluginExecutionException(name + " is not a valid id.");
            }
            return parsed;
        }

        private static string RequiredRole(IPluginExecutionContext context)
        {
            var role = Required(context, "Role").ToLowerInvariant();
            if (!ContactRoles.Contains(role))
            {
                throw new InvalidPluginExecutionException("Role must be reader or writer.");
            }
            return role;
        }
    }

    internal sealed class ContainerPermission
    {
        public string Id { get; set; }
        public List<string> Roles { get; set; }
        public string DisplayName { get; set; }
        public string Email { get; set; }
        public string UserPrincipalName { get; set; }

        public bool IsExternal
        {
            get { return (UserPrincipalName ?? "").IndexOf("#EXT#", StringComparison.OrdinalIgnoreCase) >= 0; }
        }
    }

    internal sealed class GraphSession
    {
        private const string Graph = "https://graph.microsoft.com/v1.0";
        private static readonly object TokenGate = new object();
        private static string cachedToken;
        private static string cachedFor;
        private static DateTime cachedUntilUtc = DateTime.MinValue;

        private readonly HttpClient http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        private readonly string token;
        private readonly string containerTypeId;

        private GraphSession(string token, string containerTypeId)
        {
            this.token = token;
            this.containerTypeId = containerTypeId;
        }

        public static GraphSession Create(IOrganizationService service)
        {
            var tenantId = EnvironmentValue(service, "new_tenantid");
            var clientId = EnvironmentValue(service, "new_owningappid");
            var clientSecret = EnvironmentValue(service, "new_graphclientsecret");
            var containerTypeId = EnvironmentValue(service, "new_containertypeid");
            return new GraphSession(AccessToken(tenantId, clientId, clientSecret), containerTypeId);
        }

        public string CreateContainer(string displayName, string caseId)
        {
            var payload = "{\"displayName\":\"" + JsonText.Escape(displayName)
                + "\",\"description\":\"" + JsonText.Escape(caseId)
                + "\",\"containerTypeId\":\"" + JsonText.Escape(containerTypeId) + "\"}";
            var created = JsonText.Deserialize<IdResponse>(Send(HttpMethod.Post, Graph + "/storage/fileStorage/containers", payload));
            if (created == null || string.IsNullOrWhiteSpace(created.Id))
            {
                throw new InvalidPluginExecutionException("SharePoint Embedded did not return a container id.");
            }
            return created.Id;
        }

        public void DeleteContainer(string containerId)
        {
            Send(HttpMethod.Delete, ContainerUrl(containerId), null);
        }

        public List<ContainerPermission> ListPermissions(string containerId)
        {
            var page = JsonText.Deserialize<PermissionPage>(Send(HttpMethod.Get, ContainerUrl(containerId) + "/permissions", null));
            return page == null || page.Value == null
                ? new List<ContainerPermission>()
                : page.Value.Where(p => p != null).Select(ToPermission).ToList();
        }

        public ContainerPermission GrantPermission(string containerId, string upn, string role)
        {
            var payload = "{\"roles\":[\"" + role + "\"],\"grantedToV2\":{\"user\":{\"userPrincipalName\":\"" + JsonText.Escape(upn) + "\"}}}";
            return ToPermission(JsonText.Deserialize<PermissionItem>(Send(HttpMethod.Post, ContainerUrl(containerId) + "/permissions", payload)));
        }

        public ContainerPermission UpdatePermission(string containerId, string permissionId, string role)
        {
            var payload = "{\"roles\":[\"" + role + "\"]}";
            return ToPermission(JsonText.Deserialize<PermissionItem>(
                Send(new HttpMethod("PATCH"), ContainerUrl(containerId) + "/permissions/" + Uri.EscapeDataString(permissionId), payload)));
        }

        public void DeletePermission(string containerId, string permissionId)
        {
            Send(HttpMethod.Delete, ContainerUrl(containerId) + "/permissions/" + Uri.EscapeDataString(permissionId), null);
        }

        public string FindOrInviteGuest(string email, string displayName)
        {
            var filter = Uri.EscapeDataString("mail eq '" + email.Replace("'", "''") + "'");
            var found = JsonText.Deserialize<UserPage>(Send(HttpMethod.Get, Graph + "/users?$filter=" + filter + "&$select=id,userPrincipalName", null));
            var existing = found == null || found.Value == null ? null : found.Value.FirstOrDefault();
            if (existing != null && !string.IsNullOrWhiteSpace(existing.UserPrincipalName))
            {
                return existing.UserPrincipalName;
            }

            var payload = "{\"invitedUserEmailAddress\":\"" + JsonText.Escape(email)
                + "\",\"invitedUserDisplayName\":\"" + JsonText.Escape(displayName ?? email)
                + "\",\"inviteRedirectUrl\":\"https://myapps.microsoft.com\",\"sendInvitationMessage\":true}";
            var invitation = JsonText.Deserialize<InvitationResponse>(Send(HttpMethod.Post, Graph + "/invitations", payload));
            if (invitation == null || invitation.InvitedUser == null || string.IsNullOrWhiteSpace(invitation.InvitedUser.Id))
            {
                throw new InvalidPluginExecutionException("Microsoft Entra did not create a guest for " + email + ".");
            }
            var user = JsonText.Deserialize<UserItem>(
                Send(HttpMethod.Get, Graph + "/users/" + Uri.EscapeDataString(invitation.InvitedUser.Id) + "?$select=id,userPrincipalName", null));
            if (user == null || string.IsNullOrWhiteSpace(user.UserPrincipalName))
            {
                throw new InvalidPluginExecutionException("Could not read the guest account for " + email + ".");
            }
            return user.UserPrincipalName;
        }

        private static ContainerPermission ToPermission(PermissionItem item)
        {
            var user = item == null || item.GrantedToV2 == null ? null : item.GrantedToV2.User;
            return new ContainerPermission
            {
                Id = item == null ? "" : item.Id ?? "",
                Roles = item == null || item.Roles == null ? new List<string>() : item.Roles,
                DisplayName = user == null ? "" : user.DisplayName ?? "",
                Email = user == null ? "" : user.Email ?? "",
                UserPrincipalName = user == null ? "" : user.UserPrincipalName ?? ""
            };
        }

        private static string ContainerUrl(string containerId)
        {
            return Graph + "/storage/fileStorage/containers/" + Uri.EscapeDataString(containerId);
        }

        private string Send(HttpMethod method, string url, string json)
        {
            var request = new HttpRequestMessage(method, url);
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            if (json != null)
            {
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
            }
            var response = http.SendAsync(request).GetAwaiter().GetResult();
            var body = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidPluginExecutionException("Microsoft Graph returned " + (int)response.StatusCode + ". " + Shorten(body));
            }
            return body;
        }

        private static string Shorten(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }
            var flat = value.Replace("\r", " ").Replace("\n", " ");
            return flat.Length <= 400 ? flat : flat.Substring(0, 400);
        }

        private static string EnvironmentValue(IOrganizationService service, string schemaName)
        {
            var query = new QueryExpression("environmentvariabledefinition") { ColumnSet = new ColumnSet("defaultvalue"), TopCount = 1 };
            query.Criteria.AddCondition("schemaname", ConditionOperator.Equal, schemaName);
            var values = query.AddLink("environmentvariablevalue", "environmentvariabledefinitionid", "environmentvariabledefinitionid", JoinOperator.LeftOuter);
            values.EntityAlias = "current";
            values.Columns = new ColumnSet("value");

            var rows = service.RetrieveMultiple(query).Entities;
            if (rows.Count == 0)
            {
                throw new InvalidPluginExecutionException("Add environment variable " + schemaName + ".");
            }
            var aliased = rows[0].GetAttributeValue<AliasedValue>("current.value");
            var current = aliased == null ? null : aliased.Value as string;
            if (!string.IsNullOrWhiteSpace(current))
            {
                return current;
            }
            var fallback = rows[0].GetAttributeValue<string>("defaultvalue");
            if (string.IsNullOrWhiteSpace(fallback))
            {
                throw new InvalidPluginExecutionException("Set a value for environment variable " + schemaName + ".");
            }
            return fallback;
        }

        private static string AccessToken(string tenantId, string clientId, string clientSecret)
        {
            var key = tenantId + "|" + clientId;
            lock (TokenGate)
            {
                if (cachedFor == key && !string.IsNullOrEmpty(cachedToken) && DateTime.UtcNow < cachedUntilUtc)
                {
                    return cachedToken;
                }
            }

            using (var client = new HttpClient())
            {
                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("client_id", clientId),
                    new KeyValuePair<string, string>("client_secret", clientSecret),
                    new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
                    new KeyValuePair<string, string>("grant_type", "client_credentials")
                });
                var response = client.PostAsync("https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token", content).GetAwaiter().GetResult();
                var body = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidPluginExecutionException("Microsoft Entra did not issue a token. " + Shorten(body));
                }
                var token = JsonText.Deserialize<TokenResponse>(body);
                if (token == null || string.IsNullOrWhiteSpace(token.AccessToken))
                {
                    throw new InvalidPluginExecutionException("Microsoft Entra did not issue a token.");
                }
                lock (TokenGate)
                {
                    cachedToken = token.AccessToken;
                    cachedFor = key;
                    cachedUntilUtc = DateTime.UtcNow.AddMinutes(10);
                }
                return token.AccessToken;
            }
        }
    }

    internal static class JsonText
    {
        public static string Escape(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }
            var builder = new StringBuilder(value.Length);
            foreach (var character in value)
            {
                switch (character)
                {
                    case '\\': builder.Append("\\\\"); break;
                    case '"': builder.Append("\\\""); break;
                    case '\r': builder.Append("\\r"); break;
                    case '\n': builder.Append("\\n"); break;
                    case '\t': builder.Append("\\t"); break;
                    default:
                        if (character < ' ')
                        {
                            builder.Append("\\u").Append(((int)character).ToString("x4"));
                        }
                        else
                        {
                            builder.Append(character);
                        }
                        break;
                }
            }
            return builder.ToString();
        }

        public static T Deserialize<T>(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return default(T);
            }
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)serializer.ReadObject(stream);
            }
        }
    }

    [DataContract]
    public class TokenResponse
    {
        [DataMember(Name = "access_token")]
        public string AccessToken { get; set; }
    }

    [DataContract]
    public class IdResponse
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }
    }

    [DataContract]
    public class PermissionPage
    {
        [DataMember(Name = "value")]
        public List<PermissionItem> Value { get; set; }
    }

    [DataContract]
    public class PermissionItem
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }

        [DataMember(Name = "roles")]
        public List<string> Roles { get; set; }

        [DataMember(Name = "grantedToV2")]
        public GrantedTo GrantedToV2 { get; set; }
    }

    [DataContract]
    public class GrantedTo
    {
        [DataMember(Name = "user")]
        public GrantedUser User { get; set; }
    }

    [DataContract]
    public class GrantedUser
    {
        [DataMember(Name = "displayName")]
        public string DisplayName { get; set; }

        [DataMember(Name = "email")]
        public string Email { get; set; }

        [DataMember(Name = "userPrincipalName")]
        public string UserPrincipalName { get; set; }
    }

    [DataContract]
    public class UserPage
    {
        [DataMember(Name = "value")]
        public List<UserItem> Value { get; set; }
    }

    [DataContract]
    public class UserItem
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }

        [DataMember(Name = "userPrincipalName")]
        public string UserPrincipalName { get; set; }
    }

    [DataContract]
    public class InvitationResponse
    {
        [DataMember(Name = "invitedUser")]
        public UserItem InvitedUser { get; set; }
    }
}
