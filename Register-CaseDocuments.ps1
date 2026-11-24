$ErrorActionPreference = "Stop"
$org = "https://mpcustomdemo2.crm19.dynamics.com"
$solutionName = "SharePoint-Embeded-Demo"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$project = Join-Path $root "CaseDocumentsPlugin"
$dll = Join-Path $project "bin\Release\net462\CaseDocumentsPlugin.dll"

function Get-DataverseToken {
    if ($script:dataverseToken) {
        return $script:dataverseToken
    }
    $raw = @(pac auth token 2>&1 | ForEach-Object { "$_" }) -join ""
    $match = [regex]::Match($raw, '[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+')
    if (-not $match.Success) {
        throw "pac auth token did not return a token. Run pac auth create against $org first."
    }
    $script:dataverseToken = $match.Value
    return $script:dataverseToken
}

function Invoke-Dataverse {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body
    )
    $token = Get-DataverseToken
    $headers = @{
        Authorization = "Bearer $token"
        "OData-MaxVersion" = "4.0"
        "OData-Version" = "4.0"
        Accept = "application/json"
        "Content-Type" = "application/json; charset=utf-8"
        Prefer = "return=representation"
        "MSCRM.SolutionUniqueName" = $solutionName
    }
    if ($null -eq $Body) {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
    }
    $json = $Body | ConvertTo-Json -Depth 8 -Compress
    return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $json
}

Write-Host "Checking the signed-in environment..."
$who = pac org who | Out-String
if ($who -notmatch "mpcustomdemo2") {
    throw "pac is not signed in to $org. Run: pac auth create --environment $org"
}

$base = "$org/api/data/v9.2"
$solutions = Invoke-Dataverse -Method Get -Uri "$base/solutions?`$select=uniquename,friendlyname"
$solution = $solutions.value | Where-Object { $_.uniquename -eq $solutionName }
if (-not $solution) {
    $similar = ($solutions.value | Where-Object { $_.uniquename -match "SharePoint" -or $_.friendlyname -match "SharePoint" } | ForEach-Object { $_.uniquename }) -join ", "
    throw "Solution $solutionName was not found. Similar solutions: $similar"
}

foreach ($name in @("new_containertypeid", "new_owningappid", "new_tenantid", "new_graphclientsecret")) {
    $found = Invoke-Dataverse -Method Get -Uri "$base/environmentvariabledefinitions?`$filter=schemaname eq '$name'&`$select=schemaname,defaultvalue&`$expand=environmentvariabledefinition_environmentvariablevalue(`$select=environmentvariablevalueid)"
    if (-not $found.value -or $found.value.Count -eq 0) {
        throw "Create environment variable $name in solution $solutionName, set its current value, then run this script again."
    }
    $values = $found.value[0].environmentvariabledefinition_environmentvariablevalue
    $hasCurrent = $values -and $values.Count -gt 0
    $hasDefault = -not [string]::IsNullOrWhiteSpace($found.value[0].defaultvalue)
    if (-not $hasCurrent -and -not $hasDefault) {
        throw "Set the current value of environment variable $name, then run this script again."
    }
}

Write-Host "Building the plugin..."
dotnet build $project -c Release
if ($LASTEXITCODE -ne 0) {
    throw "dotnet build failed. Install the .NET SDK from https://dotnet.microsoft.com/download and run this script again."
}
if (-not (Test-Path $dll)) {
    throw "Build did not produce $dll"
}

Write-Host "Publishing the plugin assembly..."
$existing = Invoke-Dataverse -Method Get -Uri "$base/pluginassemblies?`$filter=name eq 'CaseDocumentsPlugin'&`$select=pluginassemblyid"
if ($existing.value -and $existing.value.Count -gt 0) {
    pac plugin push --pluginId $existing.value[0].pluginassemblyid --pluginFile $dll --type Assembly
} else {
    pac plugin push --pluginFile $dll --type Assembly
}
if ($LASTEXITCODE -ne 0) {
    throw "pac plugin push failed"
}

$pluginType = $null
for ($attempt = 0; $attempt -lt 6; $attempt++) {
    $types = Invoke-Dataverse -Method Get -Uri "$base/plugintypes?`$filter=typename eq 'CaseDocumentsPlugin.CaseDocumentsPlugin'&`$select=plugintypeid,typename"
    if ($types.value -and $types.value.Count -gt 0) {
        $pluginType = $types.value[0]
        break
    }
    Start-Sleep -Seconds 5
}
if (-not $pluginType) {
    throw "The plugin assembly was pushed, but CaseDocumentsPlugin.CaseDocumentsPlugin is not visible yet."
}

$assembly = Invoke-Dataverse -Method Get -Uri "$base/pluginassemblies?`$filter=name eq 'CaseDocumentsPlugin'&`$select=pluginassemblyid"
if ($assembly.value -and $assembly.value.Count -gt 0) {
    $addAssembly = @{
        ComponentId = $assembly.value[0].pluginassemblyid
        ComponentType = 91
        SolutionUniqueName = $solutionName
        AddRequiredComponents = $true
    }
    try {
        Invoke-Dataverse -Method Post -Uri "$base/AddSolutionComponent" -Body $addAssembly | Out-Null
    } catch {
        Write-Host "Plugin assembly is already in the solution, or it could not be added: $($_.Exception.Message)"
    }
}

function Ensure-Action {
    param(
        [string]$UniqueName,
        [string]$DisplayName,
        [string[]]$Parameters
    )
    $found = Invoke-Dataverse -Method Get -Uri "$base/customapis?`$filter=uniquename eq '$UniqueName'&`$select=customapiid"
    if ($found.value -and $found.value.Count -gt 0) {
        $apiId = $found.value[0].customapiid
        $patch = @{ "plugintypeid@odata.bind" = "/plugintypes($($pluginType.plugintypeid))" }
        Invoke-Dataverse -Method Patch -Uri "$base/customapis($apiId)" -Body $patch | Out-Null
        Write-Host "Updated $UniqueName"
        return
    }

    $created = Invoke-Dataverse -Method Post -Uri "$base/customapis" -Body @{
        uniquename = $UniqueName
        name = $UniqueName
        displayname = $DisplayName
        description = $DisplayName
        bindingtype = 0
        isfunction = $false
        isprivate = $false
        allowedcustomprocessingsteptype = 0
        "plugintypeid@odata.bind" = "/plugintypes($($pluginType.plugintypeid))"
    }
    $apiId = $created.customapiid
    foreach ($parameter in $Parameters) {
        Invoke-Dataverse -Method Post -Uri "$base/customapirequestparameters" -Body @{
            uniquename = $parameter
            name = $parameter
            displayname = $parameter
            type = 10
            isoptional = $false
            "CustomAPIId@odata.bind" = "/customapis($apiId)"
        } | Out-Null
    }
    Invoke-Dataverse -Method Post -Uri "$base/customapiresponseproperties" -Body @{
        uniquename = "Result"
        name = "Result"
        displayname = "Result"
        type = 10
        "CustomAPIId@odata.bind" = "/customapis($apiId)"
    } | Out-Null
    Write-Host "Created $UniqueName"
}

Ensure-Action -UniqueName "new_EnsureCaseContainer" -DisplayName "Ensure case container" -Parameters @("CaseId")
Ensure-Action -UniqueName "new_ListCaseDocuments" -DisplayName "List case documents" -Parameters @("ContainerId")
Ensure-Action -UniqueName "new_PreviewCaseDocument" -DisplayName "Preview case document" -Parameters @("ContainerId", "ItemId")
Ensure-Action -UniqueName "new_DownloadCaseDocument" -DisplayName "Download case document" -Parameters @("ContainerId", "ItemId")
Ensure-Action -UniqueName "new_SearchCaseDocuments" -DisplayName "Search case documents" -Parameters @("ContainerId", "Query")

$stepName = "CaseDocumentsPlugin: Create of incident"
$steps = Invoke-Dataverse -Method Get -Uri "$base/sdkmessageprocessingsteps?`$filter=name eq '$stepName'&`$select=sdkmessageprocessingstepid"
if (-not $steps.value -or $steps.value.Count -eq 0) {
    $message = Invoke-Dataverse -Method Get -Uri "$base/sdkmessages?`$filter=name eq 'Create'&`$select=sdkmessageid"
    $messageId = $message.value[0].sdkmessageid
    $filter = Invoke-Dataverse -Method Get -Uri "$base/sdkmessagefilters?`$filter=primaryobjecttypecode eq 'incident' and _sdkmessageid_value eq $messageId&`$select=sdkmessagefilterid"
    if (-not $filter.value -or $filter.value.Count -eq 0) {
        throw "Could not find the Create message for the incident table."
    }
    $createdStep = Invoke-Dataverse -Method Post -Uri "$base/sdkmessageprocessingsteps" -Body @{
        name = $stepName
        description = "Create a SharePoint Embedded container when a case is saved."
        mode = 0
        rank = 1
        stage = 40
        supporteddeployment = 0
        asyncautodelete = $false
        "plugintypeid@odata.bind" = "/plugintypes($($pluginType.plugintypeid))"
        "sdkmessageid@odata.bind" = "/sdkmessages($messageId)"
        "sdkmessagefilterid@odata.bind" = "/sdkmessagefilters($($filter.value[0].sdkmessagefilterid))"
    }
    try {
        Invoke-Dataverse -Method Post -Uri "$base/AddSolutionComponent" -Body @{
            ComponentId = $createdStep.sdkmessageprocessingstepid
            ComponentType = 92
            SolutionUniqueName = $solutionName
            AddRequiredComponents = $true
        } | Out-Null
    } catch {
        Write-Host "Plugin step was created. Add it to the solution if it is not visible there."
    }
    Write-Host "Registered case-create step"
} else {
    Write-Host "Case-create step already exists"
}

Write-Host "Building and publishing the case documents control..."
Push-Location $root
npm run build
if ($LASTEXITCODE -ne 0) {
    throw "npm run build failed"
}
pac pcf push --publisher-prefix new
if ($LASTEXITCODE -ne 0) {
    throw "pac pcf push failed"
}
Pop-Location
Write-Host "Done. Open a saved case. The SharePoint container field is filled when the case documents tab loads."
