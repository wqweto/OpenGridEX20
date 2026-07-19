<#
.SYNOPSIS
Generates tools\common\mdProfiles.bas from the control/class stubs in src.

Parses Public Property Get/Let/Set declarations, classifies each property
(scalar/enum/variant/font/picture/object/class/collection) and merges the
curated runtime-only knowledge below. Parameterized properties are skipped
(handled as special cases in mdSnapshot.bas). Re-run after any API change.
#>
$ErrorActionPreference = "Stop"
$sSrc = "$PSScriptRoot\..\src"
$sOut = "$PSScriptRoot\common\mdProfiles.bas"

#--- collection class -> item class
$cCollections = @{
    "JSColumns" = "JSColumn"; "JSFmtConditions" = "JSFmtCondition"
    "JSGridImages" = "JSGridImage"; "JSGroups" = "JSGroup"
    "JSSortKeys" = "JSSortKey"; "JSValueList" = "JSValueItem"
    "JSDataObjectFiles" = "String"; "JSSelectedItems" = "JSSelectedItem"
    "JSFormatStyles" = "JSFormatStyle"
}
#--- classes whose entire state is runtime-only
$aRuntimeClasses = @("JSSelectedItem", "JSSelectedItems", "JSRowData", "JSDataObject", "JSDataObjectFiles", "JSRetBoolean", "JSRetVariant", "JSRetInteger")
#--- individual runtime-only properties (Class.Prop)
$aRuntimeProps = @(
    "GridEX.hWnd", "GridEX.hWndEdit", "GridEX.RowCount", "GridEX.Row", "GridEX.Col",
    "GridEX.FirstItem", "GridEX.LeftCol", "GridEX.EditMode", "GridEX.SelStart",
    "GridEX.SelLength", "GridEX.SelText", "GridEX.DataChanged", "GridEX.ErrorText",
    "GridEX.FullyLoaded", "GridEX.Recordset", "GridEX.ADORecordset", "GridEX.Redraw",
    "GridEX.SelectedItems", "GridEX.ItemCount",
    "GEXPreview.hWnd", "GEXPreview.CurrentPage", "GEXPreview.TotalPages",
    "JSColumn.Index", "JSColumn.IsGrouped", "JSColumn.SortOrder", "JSColumn.DataChanged",
    "JSColumn.DropDownControl",
    "JSColumns.Count", "JSFmtConditions.Count", "JSGridImages.Count", "JSGroups.Count",
    "JSSortKeys.Count", "JSValueList.Count", "JSValueList.VisibleCount", "JSFormatStyles.Count",
    "JSGridImages.hImageList",
    "JSGridImage.Index", "JSGroup.Index", "JSSortKey.Index", "JSValueItem.Index",
    "JSFmtCondition.Index",
    "JSPrinterProperties.DeviceName", "JSPrinterProperties.DriverName",
    "JSPrinterProperties.ClientWidth", "JSPrinterProperties.ClientHeight",
    "JSPrinterProperties.PageSetupCanceled"
)

function Get-Kind([string]$sType, [string]$sClass) {
    if ($cCollections.ContainsKey($sType)) { return "ucsPrkCollection" }
    if ($sType -like "jgex*") { return "ucsPrkEnum" }
    if ($sType -like "JS*") { return "ucsPrkClass" }
    switch ($sType) {
    "Font" { return "ucsPrkFont" }
    "Picture" { return "ucsPrkPicture" }
    "Variant" { return "ucsPrkVariant" }
    "Object" { return "ucsPrkObject" }
    default { return "ucsPrkScalar" }
    }
}

$aEntries = @()
$aSkipped = @()
$aFiles = @(Get-ChildItem "$sSrc\*.cls") + @(Get-Item "$sSrc\GridEX.ctl") + @(Get-Item "$sSrc\GEXPreview.ctl")
foreach ($oFile in $aFiles) {
    $sClass = $oFile.BaseName
    if ($sClass -eq "GridDesigner") { continue }
    $sText = [IO.File]::ReadAllText($oFile.FullName)
    $cProps = [ordered]@{}
    foreach ($oMatch in [regex]::Matches($sText, "(?m)^Public Property (Get|Let|Set)\s+(\w+)\s*\(([^)]*)\)(?:\s+As\s+([\w.]+))?")) {
        $sVerb = $oMatch.Groups[1].Value
        $sName = $oMatch.Groups[2].Value
        $sArgs = $oMatch.Groups[3].Value.Trim()
        $sType = $oMatch.Groups[4].Value
        if (-not $cProps.Contains($sName)) { $cProps[$sName] = @{ Get = $false; Write = $false; Type = ""; GetArgs = "" } }
        if ($sVerb -eq "Get") {
            $cProps[$sName].Get = $true
            $cProps[$sName].Type = $sType
            $cProps[$sName].GetArgs = $sArgs
        } else {
            $cProps[$sName].Write = $true
        }
    }
    foreach ($sName in $cProps.Keys) {
        $uProp = $cProps[$sName]
        if (-not $uProp.Get) { continue }
        if ($sName -eq "NewEnum") { continue }
        if ($uProp.GetArgs -ne "") { $aSkipped += "$sClass.$sName($($uProp.GetArgs))"; continue }
        $sKind = Get-Kind $uProp.Type $sClass
        $bRuntime = ($sClass -in $aRuntimeClasses) -or ("$sClass.$sName" -in $aRuntimeProps) -or ($sKind -eq "ucsPrkObject")
        $sItemClass = ""
        if ($sKind -eq "ucsPrkCollection") { $sItemClass = $cCollections[$uProp.Type] }
        $aEntries += [pscustomobject]@{
            Class = $sClass; Prop = $sName; Kind = $sKind; TypeName = $uProp.Type
            ItemClass = $sItemClass; CanWrite = $uProp.Write; RuntimeOnly = $bRuntime
        }
    }
}

$oSb = New-Object Text.StringBuilder
[void]$oSb.AppendLine('Attribute VB_Name = "mdProfiles"')
[void]$oSb.AppendLine("'=========================================================================")
[void]$oSb.AppendLine("'")
[void]$oSb.AppendLine("' Open GridEX 2000 Control")
[void]$oSb.AppendLine("' Property profiles for the snapshot engine")
[void]$oSb.AppendLine("'")
[void]$oSb.AppendLine("' GENERATED FILE -- DO NOT EDIT. Re-generate with tools\GenProfiles.ps1")
[void]$oSb.AppendLine("'")
[void]$oSb.AppendLine("' Skipped parameterized properties (handled in mdSnapshot.bas):")
foreach ($sSkip in $aSkipped) { [void]$oSb.AppendLine("'   $sSkip") }
[void]$oSb.AppendLine("'")
[void]$oSb.AppendLine("'=========================================================================")
[void]$oSb.AppendLine("Option Explicit")
[void]$oSb.AppendLine("")
[void]$oSb.AppendLine("Public Enum UcsPropKindEnum")
[void]$oSb.AppendLine("    ucsPrkScalar")
[void]$oSb.AppendLine("    ucsPrkEnum")
[void]$oSb.AppendLine("    ucsPrkVariant")
[void]$oSb.AppendLine("    ucsPrkFont")
[void]$oSb.AppendLine("    ucsPrkPicture")
[void]$oSb.AppendLine("    ucsPrkObject")
[void]$oSb.AppendLine("    ucsPrkClass")
[void]$oSb.AppendLine("    ucsPrkCollection")
[void]$oSb.AppendLine("End Enum")
[void]$oSb.AppendLine("")
[void]$oSb.AppendLine("Public Type UcsProfileProp")
[void]$oSb.AppendLine("    sClass                          As String")
[void]$oSb.AppendLine("    sProp                           As String")
[void]$oSb.AppendLine("    eKind                           As UcsPropKindEnum")
[void]$oSb.AppendLine("    sTypeName                       As String")
[void]$oSb.AppendLine("    sItemClass                      As String")
[void]$oSb.AppendLine("    bCanWrite                       As Boolean")
[void]$oSb.AppendLine("    bRuntimeOnly                    As Boolean")
[void]$oSb.AppendLine("End Type")
[void]$oSb.AppendLine("")
[void]$oSb.AppendLine("Private m_aProps()                  As UcsProfileProp")
[void]$oSb.AppendLine("Private m_lCount                    As Long")
[void]$oSb.AppendLine("")
[void]$oSb.AppendLine("Public Sub ProfilesInit()")
[void]$oSb.AppendLine("    If m_lCount > 0 Then")
[void]$oSb.AppendLine("        Exit Sub")
[void]$oSb.AppendLine("    End If")
[void]$oSb.AppendLine("    ReDim m_aProps(0 To $($aEntries.Count - 1)) As UcsProfileProp")
foreach ($uEntry in $aEntries) {
    $sWrite = if ($uEntry.CanWrite) { "True" } else { "False" }
    $sRuntime = if ($uEntry.RuntimeOnly) { "True" } else { "False" }
    [void]$oSb.AppendLine("    pvAdd `"$($uEntry.Class)`", `"$($uEntry.Prop)`", $($uEntry.Kind), `"$($uEntry.TypeName)`", `"$($uEntry.ItemClass)`", $sWrite, $sRuntime")
}
[void]$oSb.AppendLine("End Sub")
[void]$oSb.AppendLine("")
[void]$oSb.AppendLine("Public Function ProfileForClass(sClass As String, uProps() As UcsProfileProp) As Long")
[void]$oSb.AppendLine("    Dim lIdx            As Long")
[void]$oSb.AppendLine("")
[void]$oSb.AppendLine("    ProfilesInit")
[void]$oSb.AppendLine("    ReDim uProps(0 To m_lCount - 1) As UcsProfileProp")
[void]$oSb.AppendLine("    For lIdx = 0 To m_lCount - 1")
[void]$oSb.AppendLine("        If m_aProps(lIdx).sClass = sClass Then")
[void]$oSb.AppendLine("            uProps(ProfileForClass) = m_aProps(lIdx)")
[void]$oSb.AppendLine("            ProfileForClass = ProfileForClass + 1")
[void]$oSb.AppendLine("        End If")
[void]$oSb.AppendLine("    Next")
[void]$oSb.AppendLine("End Function")
[void]$oSb.AppendLine("")
[void]$oSb.AppendLine("Private Sub pvAdd(sClass As String, sProp As String, ByVal eKind As UcsPropKindEnum, sTypeName As String, sItemClass As String, ByVal bCanWrite As Boolean, ByVal bRuntimeOnly As Boolean)")
[void]$oSb.AppendLine("    With m_aProps(m_lCount)")
[void]$oSb.AppendLine("        .sClass = sClass")
[void]$oSb.AppendLine("        .sProp = sProp")
[void]$oSb.AppendLine("        .eKind = eKind")
[void]$oSb.AppendLine("        .sTypeName = sTypeName")
[void]$oSb.AppendLine("        .sItemClass = sItemClass")
[void]$oSb.AppendLine("        .bCanWrite = bCanWrite")
[void]$oSb.AppendLine("        .bRuntimeOnly = bRuntimeOnly")
[void]$oSb.AppendLine("    End With")
[void]$oSb.AppendLine("    m_lCount = m_lCount + 1")
[void]$oSb.AppendLine("End Sub")

New-Item -ItemType Directory -Force (Split-Path $sOut) | Out-Null
[IO.File]::WriteAllText($sOut, ($oSb.ToString() -replace "`r`n", "`n" -replace "`n", "`r`n"), (New-Object Text.UTF8Encoding $false))
"Generated $sOut : $($aEntries.Count) properties, $($aSkipped.Count) parameterized skipped"
