#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$TemplatePath = "firewall-policy-automation.xlsx",
    [string]$VbaModulePath = "vba\FirewallPolicyAutomation.bas",
    [string]$OutputPath = "dist\firewall-policy-automation.xlsm",
    [switch]$NoButtons,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

function Resolve-InputPath {
    param([string]$PathValue, [string]$Label)
    $resolved = Resolve-Path -LiteralPath $PathValue -ErrorAction SilentlyContinue
    if (-not $resolved) {
        throw "$Label not found: $PathValue"
    }
    return $resolved.Path
}

function Convert-ToAbsoluteOutputPath {
    param([string]$PathValue)
    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $PathValue))
}

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

$template = Resolve-InputPath -PathValue $TemplatePath -Label "Template workbook"
$vbaModule = Resolve-InputPath -PathValue $VbaModulePath -Label "VBA module"
$output = Convert-ToAbsoluteOutputPath -PathValue $OutputPath
$outputDirectory = Split-Path -Parent $output

if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

if ((Test-Path -LiteralPath $output) -and -not $Overwrite) {
    throw "Output already exists: $output. Pass -Overwrite to replace it."
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($template)

    $workbook.VBProject.VBComponents.Import($vbaModule) | Out-Null

    $excel.Run("'" + $workbook.Name + "'!SetupFirewallAutomationWorkbook")

    if (-not $NoButtons) {
        $settingsSheet = $workbook.Worksheets.Item("settings")
        $top = 10
        $left = 320
        $width = 190
        $height = 28

        $button = $settingsSheet.Buttons().Add($left, $top, $width, $height)
        $button.Characters().Text = "신청서 폴더 선택"
        $button.OnAction = "SelectRequestFolder"

        $button = $settingsSheet.Buttons().Add($left, $top + 36, $width, $height)
        $button.Characters().Text = "신청서 통합 실행"
        $button.OnAction = "MergeFirewallRequestFolder"

        $button = $settingsSheet.Buttons().Add($left, $top + 72, $width, $height)
        $button.Characters().Text = "샘플 신청서 생성"
        $button.OnAction = "CreateSampleRequestWorkbook"
    }

    # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    $workbook.SaveAs($output, 52)
    Write-Host "Created macro-enabled workbook: $output"
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
