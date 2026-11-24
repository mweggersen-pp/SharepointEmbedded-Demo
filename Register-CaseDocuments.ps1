$ErrorActionPreference = "Stop"
$org = "https://mpcustomdemo2.crm19.dynamics.com"
$publisherSolution = "SharePointEmbededDemo"
$assemblyFolder = "CaseDocumentsPlugin-c3a1d0e2-7b44-4f0a-9c31-6a0e5d2b91f4"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

function Invoke-Checked {
    param([string]$Command, [string[]]$Arguments)
    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Command $($Arguments -join ' ') failed"
    }
}

Write-Host "Checking the signed-in environment..."
$who = pac org who | Out-String
if ($who -notmatch "mpcustomdemo2") {
    throw "pac is not signed in to $org. Run: pac auth create --environment $org"
}

$work = Join-Path ([System.IO.Path]::GetTempPath()) ("case-documents-actions-" + (Get-Date -Format "yyyyMMddHHmmss"))
New-Item -ItemType Directory -Path $work | Out-Null

Write-Host "Building the plugin..."
$pluginProject = Join-Path (Join-Path $root "CaseDocumentsPlugin") "CaseDocumentsPlugin.csproj"
$pluginOutput = Join-Path $work "plugin"
Invoke-Checked "dotnet" @("build", $pluginProject, "-c", "Release", "-o", $pluginOutput, "--disable-build-servers")

$solutionFolder = Join-Path $work "solution"
Copy-Item (Join-Path (Join-Path $root "CaseDocumentsActions") "src") $solutionFolder -Recurse
Copy-Item (Join-Path $pluginOutput "CaseDocumentsPlugin.dll") (Join-Path (Join-Path $solutionFolder "PluginAssemblies") $assemblyFolder)

Write-Host "Reading the publisher from $publisherSolution..."
$exportZip = Join-Path $work "publisher.zip"
$exportFolder = Join-Path $work "publisher"
Invoke-Checked "pac" @("solution", "export", "--name", $publisherSolution, "--path", $exportZip, "--overwrite")
Invoke-Checked "pac" @("solution", "unpack", "--zipfile", $exportZip, "--folder", $exportFolder)

$exportedPath = Join-Path (Join-Path $exportFolder "Other") "Solution.xml"
$targetPath = Join-Path (Join-Path $solutionFolder "Other") "Solution.xml"
[xml]$exported = Get-Content -Path $exportedPath -Raw
[xml]$target = Get-Content -Path $targetPath -Raw
$publisher = $target.ImportNode($exported.ImportExportXml.SolutionManifest.Publisher, $true)
$manifest = $target.ImportExportXml.SolutionManifest
[void]$manifest.ReplaceChild($publisher, $manifest.Publisher)
$target.Save($targetPath)

Write-Host "Packing the case document actions..."
$importZip = Join-Path $work "CaseDocumentsActions.zip"
Invoke-Checked "pac" @("solution", "pack", "--zipfile", $importZip, "--folder", $solutionFolder, "--packagetype", "Unmanaged")

Write-Host "Importing the case document actions..."
Invoke-Checked "pac" @("solution", "import", "--path", $importZip, "--publish-changes", "--activate-plugins")
Write-Host "Done. New cases get a container automatically. Existing cases get one when the documents tab opens."
