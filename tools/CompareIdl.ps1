<#
.SYNOPSIS
Compares the public typelib surface of OpenGridEX20 against the original
GridEX20 control for source compatibility.

Parses OLE/COM Object Viewer .idl dumps into canonical member signatures and
reports missing/extra/different members. Ignores known-irrelevant differences:
GUIDs, auto-assigned dispids, restricted Missing* padding, helpstrings.
Compares: coclass names, enum names/values, member names, property kinds
(get/let/set), special dispids (0, -4, -512, -552), parameter attributes,
types, names and default values.

.EXAMPLE
powershell -File CompareIdl.ps1
powershell -File CompareIdl.ps1 -OldIdl ..\doc\GridEX20.idl -NewIdl ..\doc\OpenGridEX20.idl
#>
param(
    [string]$OldIdl = "$PSScriptRoot\..\doc\GridEX20.idl",
    [string]$NewIdl = "$PSScriptRoot\..\doc\OpenGridEX20.idl"
)
$ErrorActionPreference = "Stop"

function Parse-Idl([string]$sPath) {
    $sText = [IO.File]::ReadAllText($sPath)
    $sText = $sText -replace "//[^`r`n]*", ""
    $uResult = @{ Interfaces = @{}; Enums = @{}; CoClasses = @() }
    #--- coclasses
    foreach ($oMatch in [regex]::Matches($sText, "coclass\s+(\w+)\s*\{")) {
        $uResult.CoClasses += $oMatch.Groups[1].Value
    }
    #--- enums
    foreach ($oMatch in [regex]::Matches($sText, "typedef[^;{]*enum\s*\{(.*?)\}\s*(\w+);", "Singleline")) {
        $sBody = $oMatch.Groups[1].Value
        $sName = $oMatch.Groups[2].Value
        $cVals = [ordered]@{}
        foreach ($oVal in [regex]::Matches($sBody, "(\w+)\s*=\s*(-?\w+)")) {
            $lVal = 0
            $sRaw = $oVal.Groups[2].Value
            if ($sRaw -like "0x*") { $lVal = [int64][Convert]::ToUInt32($sRaw.Substring(2), 16); if ($lVal -gt 0x7FFFFFFF) { $lVal -= 0x100000000 } }
            else { $lVal = [int64]$sRaw }
            $cVals[$oVal.Groups[1].Value] = $lVal
        }
        $uResult.Enums[$sName] = $cVals
    }
    #--- interfaces and dispinterfaces
    foreach ($oMatch in [regex]::Matches($sText, "(?:^|\n)\s*(?:interface|dispinterface)\s+(\w+)(?:\s*:\s*\w+)?\s*\{(.*?)\n\s*\};", "Singleline")) {
        $sName = $oMatch.Groups[1].Value
        $sBody = $oMatch.Groups[2].Value
        $cMembers = @{}
        foreach ($oMem in [regex]::Matches($sBody, "\[([^\]]*)\]\s*(?:HRESULT|void)\s+(\w+)\s*\(([^;]*?)\);", "Singleline")) {
            $sAttrs = $oMem.Groups[1].Value
            $sMemName = $oMem.Groups[2].Value
            $sParams = $oMem.Groups[3].Value
            if ($sAttrs -match "restricted") { continue }
            $sKind = "method"
            if ($sAttrs -match "propget") { $sKind = "get" }
            elseif ($sAttrs -match "propputref") { $sKind = "set" }
            elseif ($sAttrs -match "propput") { $sKind = "let" }
            $sId = "auto"
            if ($sAttrs -match "id\((0x[0-9a-fA-F]+|\d+)\)") {
                $lId = 0
                $sRaw = $Matches[1]
                if ($sRaw -like "0x*") { $lId = [int64][Convert]::ToUInt32($sRaw.Substring(2), 16); if ($lId -gt 0x7FFFFFFF) { $lId -= 0x100000000 } }
                else { $lId = [int64]$sRaw }
                if ($lId -le 0) { $sId = [string]$lId }
            }
            #--- params: each starts with an [attrs] block
            $aParams = @()
            foreach ($oPar in [regex]::Matches($sParams, "\[([^\]]*)\]\s*([\w\* ]+?)(\w+)?\s*(?=,\s*\[|$)", "Singleline")) {
                $sPAttrs = ($oPar.Groups[1].Value -split ",\s*" | Sort-Object) -join ","
                $sPAttrs = $sPAttrs -replace "defaultvalue\(0xffffffff\)", "defaultvalue(-1)"
                $sPType = ($oPar.Groups[2].Value -replace "\s+", " ").Trim()
                $sPName = $oPar.Groups[3].Value
                $aParams += "[$sPAttrs]$sPType $sPName".Trim()
            }
            $cMembers["$sMemName/$sKind"] = "id=$sId (" + ($aParams -join "; ") + ")"
        }
        $uResult.Interfaces[$sName] = $cMembers
    }
    return $uResult
}

$uOld = Parse-Idl $OldIdl
$uNew = Parse-Idl $NewIdl
$lDiffs = 0

"orig: coclasses=$($uOld.CoClasses.Count) enums=$($uOld.Enums.Count) interfaces=$($uOld.Interfaces.Count)"
"ours: coclasses=$($uNew.CoClasses.Count) enums=$($uNew.Enums.Count) interfaces=$($uNew.Interfaces.Count)"
foreach ($sName in $uOld.CoClasses) { if ($sName -notin $uNew.CoClasses) { "MISSING coclass: $sName"; $lDiffs++ } }
foreach ($sName in $uNew.CoClasses) { if ($sName -notin $uOld.CoClasses) { "EXTRA coclass: $sName"; $lDiffs++ } }
foreach ($sEnum in $uOld.Enums.Keys) {
    if (-not $uNew.Enums.ContainsKey($sEnum)) { "MISSING enum: $sEnum"; $lDiffs++; continue }
    foreach ($sVal in $uOld.Enums[$sEnum].Keys) {
        if (-not $uNew.Enums[$sEnum].Contains($sVal)) { "MISSING value: $sEnum.$sVal"; $lDiffs++ }
        elseif ($uNew.Enums[$sEnum][$sVal] -ne $uOld.Enums[$sEnum][$sVal]) { "VALUE DIFF: $sEnum.$sVal orig=$($uOld.Enums[$sEnum][$sVal]) ours=$($uNew.Enums[$sEnum][$sVal])"; $lDiffs++ }
    }
}
foreach ($sEnum in $uNew.Enums.Keys) { if (-not $uOld.Enums.ContainsKey($sEnum)) { "EXTRA enum: $sEnum"; $lDiffs++ } }
foreach ($sIntf in ($uOld.Interfaces.Keys | Sort-Object)) {
    if ($sIntf -eq "IObjectSafety") { continue }
    if (-not $uNew.Interfaces.ContainsKey($sIntf)) { "MISSING interface: $sIntf"; $lDiffs++; continue }
    $cOld = $uOld.Interfaces[$sIntf]
    $cNew = $uNew.Interfaces[$sIntf]
    foreach ($sKey in ($cOld.Keys | Sort-Object)) {
        if (-not $cNew.ContainsKey($sKey)) { "MISSING: $sIntf.$sKey"; $lDiffs++ }
        elseif ($cNew[$sKey] -ne $cOld[$sKey]) {
            "DIFF: $sIntf.$sKey"
            "   orig: $($cOld[$sKey])"
            "   ours: $($cNew[$sKey])"
            $lDiffs++
        }
    }
    foreach ($sKey in ($cNew.Keys | Sort-Object)) {
        if (-not $cOld.ContainsKey($sKey)) { "EXTRA: $sIntf.$sKey"; $lDiffs++ }
    }
}
""
if ($lDiffs -eq 0) { "Surface check PASSED (no differences)" } else { "Surface check FAILED ($lDiffs differences)" }
exit $lDiffs
