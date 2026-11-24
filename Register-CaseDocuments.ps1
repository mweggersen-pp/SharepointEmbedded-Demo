$ErrorActionPreference = "Stop"
$org = "https://mpcustomdemo2.crm19.dynamics.com"
$publisherSolution = "SharePointEmbededDemo"
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

$work = Join-Path ([System.IO.Path]::GetTempPath()) "case-documents-actions"
if (Test-Path $work) {
    Remove-Item $work -Recurse -Force
}
New-Item -ItemType Directory -Path $work | Out-Null

Write-Host "Reading the publisher from $publisherSolution..."
$exportZip = Join-Path $work "publisher.zip"
$exportFolder = Join-Path $work "publisher"
Invoke-Checked "pac" @("solution", "export", "--name", $publisherSolution, "--path", $exportZip, "--overwrite")
Invoke-Checked "pac" @("solution", "unpack", "--zipfile", $exportZip, "--folder", $exportFolder)

foreach ($project in @("CaseDocumentsPlugin", "CaseDocumentsActions")) {
    $source = Join-Path $root $project
    $destination = Join-Path $work $project
    Get-ChildItem -Path $source -Recurse -File |
        Where-Object { $_.FullName.Substring($source.Length) -notmatch '^[\\/](bin|obj)[\\/]' } |
        ForEach-Object {
            $relative = $_.FullName.Substring($source.Length + 1)
            $copy = Join-Path $destination $relative
            New-Item -ItemType Directory -Path (Split-Path -Parent $copy) -Force | Out-Null
            Copy-Item $_.FullName $copy
        }
}

$exportedPath = Join-Path (Join-Path $exportFolder "Other") "Solution.xml"
$targetPath = Join-Path (Join-Path (Join-Path (Join-Path $work "CaseDocumentsActions") "src") "Other") "Solution.xml"
[xml]$exported = Get-Content -Path $exportedPath -Raw
[xml]$target = Get-Content -Path $targetPath -Raw
$publisher = $target.ImportNode($exported.ImportExportXml.SolutionManifest.Publisher, $true)
$manifest = $target.ImportExportXml.SolutionManifest
[void]$manifest.ReplaceChild($publisher, $manifest.Publisher)
$target.Save($targetPath)

Write-Host "Building the plugin and the case document actions..."
$solutionProject = Join-Path (Join-Path $work "CaseDocumentsActions") "CaseDocumentsActions.cdsproj"
Invoke-Checked "dotnet" @("build", $solutionProject, "-c", "Release")
$importZip = Join-Path (Join-Path (Join-Path (Join-Path $work "CaseDocumentsActions") "bin") "Release") "CaseDocumentsActions.zip"
if (-not (Test-Path $importZip)) {
    throw "The build did not produce $importZip"
}

Write-Host "Importing the case document actions..."
Invoke-Checked "pac" @("solution", "import", "--path", $importZip, "--publish-changes", "--activate-plugins")

Write-Host "Building and publishing the case documents control..."
Push-Location $root
try {
    Invoke-Checked "npm" @("run", "build")
    Invoke-Checked "pac" @("pcf", "push", "--publisher-prefix", "new")
} finally {
    Pop-Location
}
Write-Host "Done. Open a saved case and the SharePoint Embedded tab. The container is created for that case."
