using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CaseDocumentsPlugin
{
    public class CaseDocumentsPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(null);

            try
            {
                if (IsCaseCreate(context))
                {
                    if (context.Depth > 1)
                    {
                        return;
                    }

                    var target = (Entity)context.InputParameters["Target"];
                    if (target.Id == Guid.Empty)
                    {
                        return;
                    }

                    try
                    {
                        Ensure(service, tracing, context.InitiatingUserId, target.Id);
                    }
                    catch (Exception ex)
                    {
                        tracing.Trace("Container creation during case save failed: {0}", ex.Message);
                    }

                    return;
                }

                switch (context.MessageName)
                {
                    case "new_EnsureCaseContainer":
                        var caseId = RequiredGuid(context, "CaseId");
                        var containerId = Ensure(service, tracing, context.InitiatingUserId, caseId);
                        context.OutputParameters["Result"] = "{\"containerId\":\"" + JsonText.Escape(containerId) + "\"}";
                        break;
                    case "new_ListCaseDocuments":
                        context.OutputParameters["Result"] = List(service, Required(context, "ContainerId"));
                        break;
                    case "new_PreviewCaseDocument":
                        context.OutputParameters["Result"] = Preview(service, Required(context, "ContainerId"), Required(context, "ItemId"));
                        break;
                    case "new_DownloadCaseDocument":
                        context.OutputParameters["Result"] = Download(service, Required(context, "ContainerId"), Required(context, "ItemId"));
                        break;
                    case "new_SearchCaseDocuments":
                        context.OutputParameters["Result"] = Search(service, Required(context, "ContainerId"), Required(context, "Query"));
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

        private static bool IsCaseCreate(IPluginExecutionContext context)
        {
            return string.Equals(context.MessageName, "Create", StringComparison.OrdinalIgnoreCase)
                && string.Equals(context.PrimaryEntityName, "incident", StringComparison.OrdinalIgnoreCase)
                && context.InputParameters.Contains("Target")
                && context.InputParameters["Target"] is Entity;
        }

        private static string Ensure(IOrganizationService service, ITracingService tracing, Guid userId, Guid caseId)
        {
            var incident = service.Retrieve(
                "incident",
                caseId,
                new ColumnSet("title", "ticketnumber", "new_sharepointcontainer"));
            var existing = incident.GetAttributeValue<string>("new_sharepointcontainer");
            if (!string.IsNullOrWhiteSpace(existing))
            {
                return existing;
            }

            var graph = GraphSession.Create(service);
            var title = incident.GetAttributeValue<string>("title");
            var number = incident.GetAttributeValue<string>("ticketnumber");
            var displayName = (number + " " + title).Trim();
            if (displayName.Length == 0)
            {
                displayName = "Case " + caseId.ToString("N");
            }
            if (displayName.Length > 120)
            {
                displayName = displayName.Substring(0, 120);
            }

            var created = graph.CreateContainer(displayName, caseId.ToString());
            GrantCaseUser(service, tracing, graph, created, userId);

            var update = new Entity("incident", caseId);
            update["new_sharepointcontainer"] = created;
            service.Update(update);
            return created;
        }

        private static void GrantCaseUser(IOrganizationService service, ITracingService tracing, GraphSession graph, string containerId, Guid userId)
        {
            if (userId == Guid.Empty)
            {
                return;
            }

            try
            {
                var user = service.Retrieve("systemuser", userId, new ColumnSet("domainname", "internalemailaddress"));
                var upn = user.GetAttributeValue<string>("domainname");
                if (string.IsNullOrWhiteSpace(upn))
                {
                    upn = user.GetAttributeValue<string>("internalemailaddress");
                }
                if (!string.IsNullOrWhiteSpace(upn))
                {
                    graph.GrantWriter(containerId, upn);
                }
            }
            catch (Exception ex)
            {
                tracing.Trace("Container permission was not granted: {0}", ex.Message);
            }
        }

        private static string List(IOrganizationService service, string containerId)
        {
            var items = GraphSession.Create(service).ListFiles(containerId);
            var builder = new StringBuilder();
            builder.Append("{\"files\":[");
            for (var index = 0; index < items.Count; index++)
            {
                if (index > 0)
                {
                    builder.Append(',');
                }
                var item = items[index];
                builder.Append("{\"id\":\"")
                    .Append(JsonText.Escape(item.Id))
                    .Append("\",\"name\":\"")
                    .Append(JsonText.Escape(item.Name))
                    .Append("\",\"size\":")
                    .Append(item.Size)
                    .Append(",\"modifiedOn\":\"")
                    .Append(JsonText.Escape(item.ModifiedOn))
                    .Append("\"}");
            }
            builder.Append("]}");
            return builder.ToString();
        }

        private static string Preview(IOrganizationService service, string containerId, string itemId)
        {
            var url = GraphSession.Create(service).PreviewUrl(containerId, itemId);
            return "{\"previewUrl\":\"" + JsonText.Escape(url) + "\"}";
        }

        private static string Download(IOrganizationService service, string containerId, string itemId)
        {
            var url = GraphSession.Create(service).DownloadUrl(containerId, itemId);
            return "{\"downloadUrl\":\"" + JsonText.Escape(url) + "\"}";
        }

        private static string Search(IOrganizationService service, string containerId, string query)
        {
            var hits = GraphSession.Create(service).Search(containerId, query);
            var builder = new StringBuilder();
            builder.Append("{\"hits\":[");
            for (var index = 0; index < hits.Count; index++)
            {
                if (index > 0)
                {
                    builder.Append(',');
                }
                var hit = hits[index];
                builder.Append("{\"itemId\":\"")
                    .Append(JsonText.Escape(hit.ItemId))
                    .Append("\",\"name\":\"")
                    .Append(JsonText.Escape(hit.Name))
                    .Append("\",\"webUrl\":\"")
                    .Append(JsonText.Escape(hit.WebUrl))
                    .Append("\",\"text\":\"")
                    .Append(JsonText.Escape(hit.Text))
                    .Append("\"}");
            }
            builder.Append("]}");
            return builder.ToString();
        }

        private static string Required(IPluginExecutionContext context, string name)
        {
            if (!context.InputParameters.Contains(name) || context.InputParameters[name] == null)
            {
                throw new InvalidPluginExecutionException(name + " is required.");
            }
            var value = context.InputParameters[name].ToString();
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidPluginExecutionException(name + " is required.");
            }
            return value;
        }

        private static Guid RequiredGuid(IPluginExecutionContext context, string name)
        {
            var value = Required(context, name).Trim('{', '}');
            Guid parsed;
            if (!Guid.TryParse(value, out parsed))
            {
                throw new InvalidPluginExecutionException(name + " is not a valid id.");
            }
            return parsed;
        }
    }

    internal sealed class GraphSession
    {
        private static readonly object TokenGate = new object();
        private static string cachedToken;
        private static DateTime cachedUntilUtc = DateTime.MinValue;

        private readonly HttpClient http = new HttpClient();
        private readonly string token;
        private readonly string containerTypeId;

        private GraphSession(string token, string containerTypeId)
        {
            this.token = token;
            this.containerTypeId = containerTypeId;
            http.Timeout = TimeSpan.FromSeconds(60);
        }

        public string CreateContainer(string displayName, string caseId)
        {
            var payload = "{\"displayName\":\"" + JsonText.Escape(displayName)
                + "\",\"description\":\"" + JsonText.Escape(caseId)
                + "\",\"containerTypeId\":\"" + JsonText.Escape(containerTypeId)
                + "\"}";
            var created = Send<ContainerCreated>(HttpMethod.Post, "https://graph.microsoft.com/v1.0/storage/fileStorage/containers", payload);
            if (created == null || string.IsNullOrWhiteSpace(created.Id))
            {
                throw new InvalidPluginExecutionException("SharePoint Embedded did not return a container id.");
            }
            return created.Id;
        }

        public void GrantWriter(string containerId, string userPrincipalName)
        {
            var payload = "{\"roles\":[\"writer\"],\"grantedToV2\":{\"user\":{\"userPrincipalName\":\""
                + JsonText.Escape(userPrincipalName) + "\"}}}";
            Send<EmptyResponse>(
                HttpMethod.Post,
                "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/" + Uri.EscapeDataString(containerId) + "/permissions",
                payload);
        }

        public List<CaseFileItem> ListFiles(string containerId)
        {
            var page = Send<DriveItemPage>(
                HttpMethod.Get,
                "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/"
                    + Uri.EscapeDataString(containerId)
                    + "/drive/root/children?$select=id,name,size,lastModifiedDateTime",
                null);
            var files = new List<CaseFileItem>();
            if (page == null || page.Value == null)
            {
                return files;
            }
            foreach (var item in page.Value)
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Id))
                {
                    continue;
                }
                files.Add(new CaseFileItem
                {
                    Id = item.Id,
                    Name = item.Name ?? "",
                    Size = item.Size,
                    ModifiedOn = item.ModifiedOn ?? ""
                });
            }
            return files;
        }

        public string PreviewUrl(string containerId, string itemId)
        {
            var preview = Send<PreviewResponse>(
                HttpMethod.Post,
                ItemUrl(containerId, itemId) + "/preview",
                "{\"zoom\":1}");
            if (preview == null || string.IsNullOrWhiteSpace(preview.GetUrl))
            {
                throw new InvalidPluginExecutionException("Preview did not return a URL.");
            }
            var url = preview.GetUrl;
            return url.IndexOf("nb=true", StringComparison.OrdinalIgnoreCase) >= 0
                ? url
                : url + (url.Contains("?") ? "&nb=true" : "?nb=true");
        }

        public string DownloadUrl(string containerId, string itemId)
        {
            var raw = SendRaw(HttpMethod.Get, ItemUrl(containerId, itemId), null);
            var url = JsonText.ReadString(raw, "@microsoft.graph.downloadUrl");
            if (string.IsNullOrWhiteSpace(url))
            {
                throw new InvalidPluginExecutionException("Download did not return a URL.");
            }
            return url;
        }

        public List<SearchHitItem> Search(string containerId, string query)
        {
            var safeQuery = (query ?? "").Replace("\"", "").Trim();
            var payload = "{\"requests\":[{\"entityTypes\":[\"driveItem\"],\"query\":{\"queryString\":\""
                + JsonText.Escape(safeQuery + " AND ContainerId:" + containerId)
                + "\"},\"sharePointOneDriveOptions\":{\"includeHiddenContent\":true},\"from\":0,\"size\":25,\"fields\":[\"id\",\"name\",\"webUrl\"]}]}";
            var response = Send<SearchResponse>(HttpMethod.Post, "https://graph.microsoft.com/beta/search/query", payload);
            var hits = new List<SearchHitItem>();
            if (response == null || response.Value == null)
            {
                return hits;
            }
            foreach (var bucket in response.Value)
            {
                if (bucket == null || bucket.HitsContainers == null)
                {
                    continue;
                }
                foreach (var container in bucket.HitsContainers)
                {
                    if (container == null || container.Hits == null)
                    {
                        continue;
                    }
                    foreach (var hit in container.Hits)
                    {
                        if (hit == null)
                        {
                            continue;
                        }
                        hits.Add(new SearchHitItem
                        {
                            ItemId = First(hit.HitId, hit.Resource == null ? null : hit.Resource.Id),
                            Name = First(hit.Resource == null ? null : hit.Resource.Name, hit.HitId),
                            WebUrl = hit.Resource == null ? "" : (hit.Resource.WebUrl ?? ""),
                            Text = hit.Summary ?? ""
                        });
                    }
                }
            }
            return hits;
        }

        public static GraphSession Create(IOrganizationService service)
        {
            var tenantId = EnvironmentValue(service, CaseDocumentsPluginTenant.Tenant);
            var clientId = EnvironmentValue(service, CaseDocumentsPluginTenant.App);
            var clientSecret = EnvironmentValue(service, CaseDocumentsPluginTenant.Secret);
            var containerTypeId = EnvironmentValue(service, CaseDocumentsPluginTenant.ContainerType);
            return new GraphSession(AccessToken(tenantId, clientId, clientSecret), containerTypeId);
        }

        private static string ItemUrl(string containerId, string itemId)
        {
            return "https://graph.microsoft.com/v1.0/storage/fileStorage/containers/"
                + Uri.EscapeDataString(containerId)
                + "/drive/items/"
                + Uri.EscapeDataString(itemId);
        }

        private static string First(string preferred, string fallback)
        {
            return string.IsNullOrWhiteSpace(preferred) ? (fallback ?? "") : preferred;
        }

        private T Send<T>(HttpMethod method, string url, string json)
        {
            var raw = SendRaw(method, url, json);
            if (typeof(T) == typeof(EmptyResponse) || string.IsNullOrWhiteSpace(raw))
            {
                return default(T);
            }
            return JsonText.Deserialize<T>(raw);
        }

        private string SendRaw(HttpMethod method, string url, string json)
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
                throw new InvalidPluginExecutionException(
                    "Microsoft Graph returned " + (int)response.StatusCode + ". " + Trim(body));
            }
            return body;
        }

        private static string Trim(string value)
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
            var query = new QueryExpression("environmentvariabledefinition")
            {
                ColumnSet = new ColumnSet("defaultvalue"),
                TopCount = 1
            };
            query.Criteria.AddCondition("schemaname", ConditionOperator.Equal, schemaName);
            var values = query.AddLink(
                "environmentvariablevalue",
                "environmentvariabledefinitionid",
                "environmentvariabledefinitionid",
                JoinOperator.LeftOuter);
            values.EntityAlias = "current";
            values.Columns = new ColumnSet("value");

            var results = service.RetrieveMultiple(query);
            var row = results.Entities.Count == 0 ? null : results.Entities[0];
            if (row == null)
            {
                throw new InvalidPluginExecutionException("Add environment variable " + schemaName + ".");
            }

            if (row.Contains("current.value"))
            {
                var aliased = row.GetAttributeValue<AliasedValue>("current.value");
                var current = aliased == null ? null : aliased.Value as string;
                if (!string.IsNullOrWhiteSpace(current))
                {
                    return current;
                }
            }

            var fallback = row.GetAttributeValue<string>("defaultvalue");
            if (string.IsNullOrWhiteSpace(fallback))
            {
                throw new InvalidPluginExecutionException("Set a value for environment variable " + schemaName + ".");
            }
            return fallback;
        }

        private static string AccessToken(string tenantId, string clientId, string clientSecret)
        {
            lock (TokenGate)
            {
                if (!string.IsNullOrEmpty(cachedToken) && DateTime.UtcNow < cachedUntilUtc)
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
                var response = client.PostAsync(
                    "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token",
                    content).GetAwaiter().GetResult();
                var body = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidPluginExecutionException("Microsoft Entra did not issue a token. " + Trim(body));
                }
                var token = JsonText.Deserialize<TokenResponse>(body);
                if (token == null || string.IsNullOrWhiteSpace(token.AccessToken))
                {
                    throw new InvalidPluginExecutionException("Microsoft Entra did not issue a token.");
                }
                lock (TokenGate)
                {
                    cachedToken = token.AccessToken;
                    cachedUntilUtc = DateTime.UtcNow.AddMinutes(10);
                }
                return token.AccessToken;
            }
        }
    }

    internal static class CaseDocumentsPluginTenant
    {
        public const string Tenant = "new_tenantid";
        public const string App = "new_owningappid";
        public const string Secret = "new_graphclientsecret";
        public const string ContainerType = "new_containertypeid";
    }

    internal static class JsonText
    {
        public static string Escape(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }
            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n");
        }

        public static T Deserialize<T>(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)serializer.ReadObject(stream);
            }
        }

        public static string ReadString(string json, string name)
        {
            var marker = "\"" + name + "\":\"";
            var start = json.IndexOf(marker, StringComparison.Ordinal);
            if (start < 0)
            {
                return "";
            }
            start += marker.Length;
            var builder = new StringBuilder();
            for (var index = start; index < json.Length; index++)
            {
                var character = json[index];
                if (character == '\\' && index + 1 < json.Length)
                {
                    builder.Append(json[index + 1]);
                    index++;
                    continue;
                }
                if (character == '"')
                {
                    break;
                }
                builder.Append(character);
            }
            return builder.ToString();
        }
    }

    internal sealed class CaseFileItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public long Size { get; set; }
        public string ModifiedOn { get; set; }
    }

    internal sealed class SearchHitItem
    {
        public string ItemId { get; set; }
        public string Name { get; set; }
        public string WebUrl { get; set; }
        public string Text { get; set; }
    }

    [DataContract]
    public class TokenResponse
    {
        [DataMember(Name = "access_token")]
        public string AccessToken { get; set; }
    }

    [DataContract]
    public class ContainerCreated
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }
    }

    [DataContract]
    public class EmptyResponse
    {
    }

    [DataContract]
    public class DriveItemPage
    {
        [DataMember(Name = "value")]
        public List<DriveItem> Value { get; set; }
    }

    [DataContract]
    public class DriveItem
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }

        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "size")]
        public long Size { get; set; }

        [DataMember(Name = "lastModifiedDateTime")]
        public string ModifiedOn { get; set; }
    }

    [DataContract]
    public class PreviewResponse
    {
        [DataMember(Name = "getUrl")]
        public string GetUrl { get; set; }
    }

    [DataContract]
    public class DownloadItem
    {
        [DataMember(Name = "@microsoft.graph.downloadUrl")]
        public string DownloadUrl { get; set; }
    }

    [DataContract]
    public class SearchResponse
    {
        [DataMember(Name = "value")]
        public List<SearchBucket> Value { get; set; }
    }

    [DataContract]
    public class SearchBucket
    {
        [DataMember(Name = "hitsContainers")]
        public List<SearchHitsContainer> HitsContainers { get; set; }
    }

    [DataContract]
    public class SearchHitsContainer
    {
        [DataMember(Name = "hits")]
        public List<SearchHit> Hits { get; set; }
    }

    [DataContract]
    public class SearchHit
    {
        [DataMember(Name = "hitId")]
        public string HitId { get; set; }

        [DataMember(Name = "summary")]
        public string Summary { get; set; }

        [DataMember(Name = "resource")]
        public SearchResource Resource { get; set; }
    }

    [DataContract]
    public class SearchResource
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }

        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "webUrl")]
        public string WebUrl { get; set; }
    }
}
