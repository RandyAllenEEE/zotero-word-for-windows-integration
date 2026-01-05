
$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot

if (!(Test-Path "$repoRoot\build\template\Zotero.dotm")) {
    Write-Host "Error: Could not find build/template/Zotero.dotm in script directory: $repoRoot" -ForegroundColor Red
    exit 1
}

if ($null -eq $repoRoot) {
    Write-Host "Error: Could not locate repository root." -ForegroundColor Red
    exit 1
}

$installDotm = "$repoRoot\install\Zotero.dotm"
$outputDotm = "$repoRoot\Zotero_New.dotm"
$sourceVbaDir = "$repoRoot\build\template\Zotero.dotm\word\vbaProject.bin"
$customUiXml = "$repoRoot\build\template\Zotero.dotm\customUI\customUI.xml"

Write-Host "Repo Root: $repoRoot"
Write-Host "Source Template: $installDotm"

# 1. Copy original template to output
Copy-Item -Path $installDotm -Destination $outputDotm -Force

# 2. Replace CustomUI XML using Zip
Write-Host "Updating CustomUI XML..."
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipMode = [System.IO.Compression.ZipArchiveMode]::Update
$zip = [System.IO.Compression.ZipFile]::Open($outputDotm, $zipMode)
$entry = $zip.GetEntry("customUI/customUI.xml")
if ($entry) { $entry.Delete() }
$entry = $zip.CreateEntry("customUI/customUI.xml")
$stream = $entry.Open()
$bytes = [System.IO.File]::ReadAllBytes($customUiXml)
$stream.Write($bytes, 0, $bytes.Length)
$stream.Close()
$zip.Dispose()

# 3. Update Macros using Word Automation
Write-Host "Updating VBA Macros via Word Automation..."
$word = New-Object -ComObject Word.Application
$word.Visible = $true # Visible to ensure it runs correctly and user sees it
$doc = $null

try {
    $doc = $word.Documents.Open($outputDotm)
    
    # Remove existing modules
    $vbProject = $doc.VBProject
    $toRemove = @("Zotero", "ZoteroRibbon")
    
    foreach ($comp in $vbProject.VBComponents) {
        if ($toRemove -contains $comp.Name) {
            Write-Host "Removing module: $($comp.Name)"
            $vbProject.VBComponents.Remove($comp)
        }
    }
    
    # Import new modules
    $filesToImport = @(
        "$sourceVbaDir\Zotero.bas",
        "$sourceVbaDir\ZoteroRibbon.bas"
    )
    
    foreach ($file in $filesToImport) {
        Write-Host "Importing: $file"
        $vbProject.VBComponents.Import($file)
    }
    
    $doc.Save()
    Write-Host "Success! Created: $outputDotm" -ForegroundColor Green
}
catch {
    Write-Host "Error during Word Automation: $_" -ForegroundColor Red
    Write-Host "Ensure you have enabled 'Trust access to the VBA project object model' in Word Trust Center Settings." -ForegroundColor Yellow
}
finally {
    if ($doc) { $doc.Close() }
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
