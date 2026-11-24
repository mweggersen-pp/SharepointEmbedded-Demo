$ErrorActionPreference = "Stop"
$org = "https://mpcustomdemo2.crm19.dynamics.com"
$solutionName = "SharePoint-Embeded-Demo"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$project = Join-Path $root "CaseDocumentsPlugin"
$dll = Join-Path $project "bin\Release\net462\CaseDocumentsPlugin.dll"
$assemblyId = "c3a1d0e2-7b44-4f0a-9c31-6a0e5d2b91f4"
$pluginTypeId = "b7e2a1c4-5d33-4a18-8f60-2c9b0e4d71aa"
$assemblyFullName = "CaseDocumentsPlugin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
$typeName = "CaseDocumentsPlugin.CaseDocumentsPlugin, CaseDocumentsPlugin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"

function Invoke-Pac {
    param([string[]]$Arguments)
    & pac @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "pac $($Arguments -join ' ') failed"
    }
}

function Write-Utf8NoBom {
    param([string]$Path, [string]$Value)
    $encoding = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($Path, $Value.TrimStart(), $encoding)
}

Write-Host "Checking the signed-in environment..."
$who = pac org who | Out-String
if ($who -notmatch "mpcustomdemo2") {
    throw "pac is not signed in to $org. Run: pac auth create --environment $org"
}

Write-Host "Building the plugin..."
dotnet build $project -c Release
if ($LASTEXITCODE -ne 0) {
    throw "dotnet build failed. Install the .NET SDK from https://dotnet.microsoft.com/download and run this script again."
}
if (-not (Test-Path $dll)) {
    throw "Build did not produce $dll"
}

$work = Join-Path $env:TEMP "case-documents-actions"
if (Test-Path $work) {
    Remove-Item $work -Recurse -Force
}
New-Item -ItemType Directory -Path $work | Out-Null
$exportZip = Join-Path $work "existing.zip"
$unpack = Join-Path $work "existing"
$pack = Join-Path $work "pack"
$importZip = Join-Path $work "CaseDocumentsActions.zip"

Write-Host "Reading the publisher from $solutionName..."
pac solution export --name $solutionName --path $exportZip --overwrite
if ($LASTEXITCODE -ne 0) {
    Write-Host "Solutions in this environment:"
    pac solution list
    throw "Could not export solution $solutionName. If the unique name above is different, tell me that name."
}
Invoke-Pac -Arguments @("solution", "unpack", "--zipfile", $exportZip, "--folder", $unpack)
$solutionXmlPath = Join-Path $unpack "Other\Solution.xml"
if (-not (Test-Path $solutionXmlPath)) {
    $solutionXmlPath = Join-Path $unpack "src\Other\Solution.xml"
}
[xml]$exported = Get-Content -Path $solutionXmlPath
$publisherXml = $exported.ImportExportXml.SolutionManifest.Publisher.OuterXml

$actions = @(
    @{ Name = "new_EnsureCaseContainer"; Display = "Ensure case container"; Parameters = @("CaseId") },
    @{ Name = "new_ListCaseDocuments"; Display = "List case documents"; Parameters = @("ContainerId") },
    @{ Name = "new_PreviewCaseDocument"; Display = "Preview case document"; Parameters = @("ContainerId", "ItemId") },
    @{ Name = "new_DownloadCaseDocument"; Display = "Download case document"; Parameters = @("ContainerId", "ItemId") },
    @{ Name = "new_SearchCaseDocuments"; Display = "Search case documents"; Parameters = @("ContainerId", "Query") }
)

New-Item -ItemType Directory -Path (Join-Path $pack "Other") | Out-Null
$pluginFolder = Join-Path $pack "PluginAssemblies\CaseDocumentsPlugin-$assemblyId"
New-Item -ItemType Directory -Path $pluginFolder | Out-Null
Copy-Item $dll (Join-Path $pluginFolder "CaseDocumentsPlugin.dll")

foreach ($action in $actions) {
    $apiFolder = Join-Path $pack "customapis\$($action.Name)"
    New-Item -ItemType Directory -Path (Join-Path $apiFolder "customapirequestparameters") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $apiFolder "customapiresponseproperties") -Force | Out-Null
    @"
<customapi uniquename="$($action.Name)">
  <allowedcustomprocessingsteptype>0</allowedcustomprocessingsteptype>
  <bindingtype>0</bindingtype>
  <boundentitylogicalname />
  <description default="$($action.Display)">
    <label description="$($action.Display)" languagecode="1033" />
  </description>
  <displayname default="$($action.Display)">
    <label description="$($action.Display)" languagecode="1033" />
  </displayname>
  <iscustomizable>1</iscustomizable>
  <executeprivilegename />
  <isfunction>0</isfunction>
  <isprivate>0</isprivate>
  <name>$($action.Name)</name>
  <plugintypeid>
    <plugintypeexportkey>{$pluginTypeId}</plugintypeexportkey>
  </plugintypeid>
  <workflowsdkstepenabled>0</workflowsdkstepenabled>
</customapi>
"@ | ForEach-Object { Write-Utf8NoBom -Path (Join-Path $apiFolder "customapi.xml") -Value $_ }
    foreach ($parameter in $action.Parameters) {
        $parameterFolder = Join-Path $apiFolder "customapirequestparameters\$parameter"
        New-Item -ItemType Directory -Path $parameterFolder -Force | Out-Null
        @"
<customapirequestparameter uniquename="$parameter">
  <description default="$parameter">
    <label description="$parameter" languagecode="1033" />
  </description>
  <displayname default="$parameter">
    <label description="$parameter" languagecode="1033" />
  </displayname>
  <iscustomizable>1</iscustomizable>
  <isoptional>0</isoptional>
  <logicalentityname />
  <name>$($action.Name).$parameter</name>
  <type>10</type>
</customapirequestparameter>
"@ | ForEach-Object { Write-Utf8NoBom -Path (Join-Path $parameterFolder "customapirequestparameter.xml") -Value $_ }
    }
    $resultFolder = Join-Path $apiFolder "customapiresponseproperties\Result"
    New-Item -ItemType Directory -Path $resultFolder -Force | Out-Null
    @"
<customapiresponseproperty uniquename="Result">
  <description default="Result">
    <label description="Result" languagecode="1033" />
  </description>
  <displayname default="Result">
    <label description="Result" languagecode="1033" />
  </displayname>
  <iscustomizable>1</iscustomizable>
  <logicalentityname />
  <name>$($action.Name).Result</name>
  <type>10</type>
</customapiresponseproperty>
"@ | ForEach-Object { Write-Utf8NoBom -Path (Join-Path $resultFolder "customapiresponseproperty.xml") -Value $_ }
}

@"
<?xml version="1.0" encoding="utf-8"?>
<EntityRelationships xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
"@ | ForEach-Object { Write-Utf8NoBom -Path (Join-Path $pack "Other\Relationships.xml") -Value $_ }

@"
<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Entities />
  <Roles />
  <Workflows />
  <FieldSecurityProfiles />
  <Templates />
  <EntityMaps />
  <EntityRelationships />
  <OrganizationSettings />
  <optionsets />
  <CustomControls />
  <EntityDataProviders />
  <PluginAssemblies>
    <PluginAssembly FullName="$assemblyFullName" PluginAssemblyId="{$assemblyId}" CustomizationLevel="1">
      <IsolationMode>2</IsolationMode>
      <SourceType>0</SourceType>
      <IntroducedVersion>1.0</IntroducedVersion>
      <FileName>/PluginAssemblies/CaseDocumentsPlugin-$assemblyId/CaseDocumentsPlugin.dll</FileName>
      <PluginTypes>
        <PluginType AssemblyQualifiedName="$typeName" PluginTypeId="{$pluginTypeId}" Name="CaseDocumentsPlugin.CaseDocumentsPlugin" FriendlyName="CaseDocumentsPlugin">
          <PluginTypeExportKey>{$pluginTypeId}</PluginTypeExportKey>
        </PluginType>
      </PluginTypes>
    </PluginAssembly>
  </PluginAssemblies>
  <Languages>
    <Language>1033</Language>
    <Language>1044</Language>
  </Languages>
</ImportExportXml>
"@ | ForEach-Object { Write-Utf8NoBom -Path (Join-Path $pack "Other\Customizations.xml") -Value $_ }

$rootComponents = New-Object System.Text.StringBuilder
[void]$rootComponents.AppendLine("      <RootComponent type=`"91`" id=`"{$assemblyId}`" behavior=`"0`" />")
[void]$rootComponents.AppendLine("      <RootComponent type=`"90`" id=`"{$pluginTypeId}`" behavior=`"0`" />")
foreach ($action in $actions) {
    [void]$rootComponents.AppendLine("      <RootComponent type=`"10034`" schemaName=`"$($action.Name)`" behavior=`"0`" />")
}

@"
<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml version="9.2.25064.132" SolutionPackageVersion="9.2" languagecode="1033" generatedBy="CrmLive" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <SolutionManifest>
    <UniqueName>CaseDocumentsActions</UniqueName>
    <LocalizedNames>
      <LocalizedName description="Case documents actions" languagecode="1033" />
      <LocalizedName description="Case documents actions" languagecode="1044" />
    </LocalizedNames>
    <Descriptions />
    <Version>1.0.1.0</Version>
    <Managed>0</Managed>
    $publisherXml
    <RootComponents>
$($rootComponents.ToString())    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>
"@ | ForEach-Object { Write-Utf8NoBom -Path (Join-Path $pack "Other\Solution.xml") -Value $_ }

Write-Host "Packing and importing the case document actions..."
Invoke-Pac -Arguments @("solution", "pack", "--zipfile", $importZip, "--folder", $pack, "--packagetype", "Unmanaged")
Invoke-Pac -Arguments @("solution", "import", "--path", $importZip, "--publish-changes", "--activate-plugins", "--force-overwrite")

Write-Host "Building and publishing the case documents control..."
Push-Location $root
npm run build
if ($LASTEXITCODE -ne 0) {
    Pop-Location
    throw "npm run build failed"
}
pac pcf push --publisher-prefix new
if ($LASTEXITCODE -ne 0) {
    Pop-Location
    throw "pac pcf push failed"
}
Pop-Location
Write-Host "Done. Open a saved case and the SharePoint Embedded tab. The container is created for that case."
