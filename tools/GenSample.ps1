<#
.SYNOPSIS
Generates VB6 setup code from a recorded design-time snapshot.

The third consumer of the M1 snapshot engine (after the test harness and the
export add-in): turns a snapshot recorded from the *original* control into
plain VB6 that configures *our* control, so a Janus sample can be ported by
re-pointing the reference and calling the generated Sub. Emitting readable
code rather than importing JSON at runtime is the point -- a sample has to
look like ordinary code written against the control, which is what source
compatibility means.

Property kinds and write/runtime flags come from tools\common\mdProfiles.bas,
enum constant names from the Public Enum blocks in src\GridEX.ctl, so the
output reads `jgexUnbound` rather than `99`.

.EXAMPLE
powershell -File GenSample.ps1 -Snapshot ..\test\snapshots\015-Unbound-1_Form1_GridEX1.json
powershell -File GenSample.ps1 -Snapshot ..\test\snapshots\016-Unbound-2_Form1_GridEX1.json -Out mdSetup.bas
#>
param(
    [Parameter(Mandatory = $true)][string]$Snapshot,
    [string]$Out,
    [string]$ControlName,
    [string]$SubName,
    [string]$Profiles,
    [string]$EnumSource
)
$ErrorActionPreference = "Stop"

#--- $PSScriptRoot is not populated in param() defaults under -File
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Profiles) { $Profiles = Join-Path $root "common\mdProfiles.bas" }
if (-not $EnumSource) { $EnumSource = Join-Path $root "..\src\GridEX.ctl" }

#--- property profiles: class -> prop -> kind/type/canWrite/runtimeOnly
$profs = @{}
foreach ($line in Get-Content $Profiles) {
    if ($line -match '^\s*pvAdd\s+"([^"]+)",\s*"([^"]+)",\s*(\w+),\s*"([^"]*)",\s*"([^"]*)",\s*(True|False),\s*(True|False)') {
        $cls = $Matches[1]
        if (-not $profs.ContainsKey($cls)) { $profs[$cls] = @{} }
        $profs[$cls][$Matches[2]] = [pscustomobject]@{
            Kind = $Matches[3]; Type = $Matches[4]; ItemClass = $Matches[5]
            CanWrite = ($Matches[6] -eq 'True'); RuntimeOnly = ($Matches[7] -eq 'True')
        }
    }
}

#--- enum value -> constant name, so numbers come out as names
$enums = @{}
$curEnum = $null; $next = 0
foreach ($line in Get-Content $EnumSource) {
    if ($line -match '^\s*Public Enum\s+(\w+)') { $curEnum = $Matches[1]; $enums[$curEnum] = @{}; $next = 0; continue }
    if ($line -match '^\s*End Enum') { $curEnum = $null; continue }
    if ($curEnum -and $line -match '^\s*(\w+)\s*=\s*(-?&H[0-9A-Fa-f]+&?|-?\d+)') {
        $v = $Matches[2] -replace '&$',''
        $val = if ($v -match '^-?&H') { [Convert]::ToInt64(($v -replace '&H',''), 16) } else { [int64]$v }
        if (-not $enums[$curEnum].ContainsKey($val)) { $enums[$curEnum][$val] = $Matches[1] }
        $next = $val + 1
    } elseif ($curEnum -and $line -match '^\s*(\w+)\s*$') {
        if (-not $enums[$curEnum].ContainsKey($next)) { $enums[$curEnum][$next] = $Matches[1] }
        $next = $next + 1
    }
}

function Format-Literal($value, $prof) {
    if ($null -eq $value) { return $null }
    if ($prof.Kind -eq 'ucsPrkEnum' -and $enums.ContainsKey($prof.Type)) {
        $n = $enums[$prof.Type][[int64]$value]
        if ($n) { return $n }
    }
    switch ($prof.Type) {
        'String'    { return '"' + ($value -replace '"','""') + '"' }
        'Boolean'   { return $(if ($value) { 'True' } else { 'False' }) }
        'OLE_COLOR' {
            $v = [int64]$value
            #--- 0xFFFFFFFF parses as Int32 -1, so the mask has to be typed
            if ($v -lt 0) { return '&H' + ('{0:X8}' -f ($v -band 0xFFFFFFFFL)) + '&' }
            return '&H' + ('{0:X6}' -f $v) + '&'
        }
    }
    if ($prof.Type -eq 'Variant' -and $value -is [string]) { return '"' + ($value -replace '"','""') + '"' }
    if ($value -is [bool]) { return $(if ($value) { 'True' } else { 'False' }) }
    return "$value"
}

$script:skipped = @()

#--- one object's writable state as VB6 lines, mirroring what ImportObject in
#--- tools\common\mdImport.bas assigns: scalars, fonts and -- recursively --
#--- read-only sub-objects such as PrinterProperties or a style's FormatStyle
function Emit-Object($cls, $obj, $pad, $exclude) {
    $lines = @()
    foreach ($name in ($obj.PSObject.Properties.Name | Sort-Object)) {
        if ($name -eq 'items') { continue }
        if ($exclude -and $name -in $exclude) { continue }
        #--- $errors is the snapshot's own record of unreadable properties
        if ($name.StartsWith('$')) { continue }
        #--- parameterized properties snapshot as an index -> value map
        if ($cls -eq 'JSPrinterProperties' -and $name -in @('HeaderString','FooterString')) {
            foreach ($pos in ($obj.$name.PSObject.Properties.Name | Sort-Object)) {
                $pv = $obj.$name.$pos
                if ($null -eq $pv) { continue }
                $lines += "$pad.$name($pos) = """ + ($pv -replace '"','""') + '"'
            }
            continue
        }
        $prof = $profs[$cls][$name]
        if (-not $prof) { $script:skipped += "$cls.$name (not in profiles)"; continue }
        if ($prof.RuntimeOnly) { continue }
        $val = $obj.$name
        if ($prof.Kind -eq 'ucsPrkFont') {
            #--- fonts drive row and header heights, so a sample that omits
            #--- them renders differently from the original -- emit rather
            #--- than skip
            if ($val) {
                $lines += "${pad}With .$name"
                foreach ($fp in @('Name','Size','Bold','Italic','Underline','Strikethrough','Charset')) {
                    $fv = $val.$fp
                    if ($null -eq $fv) { continue }
                    if ($fp -eq 'Name') { $lines += "$pad    .Name = """ + $fv + '"' }
                    elseif ($fv -is [bool]) { $lines += "$pad    .$fp = $(if ($fv) { 'True' } else { 'False' })" }
                    else { $lines += "$pad    .$fp = $fv" }
                }
                $lines += "${pad}End With"
            }
        } elseif ($prof.Kind -eq 'ucsPrkClass') {
            if ($val) {
                $sub = Emit-Object $prof.Type $val "$pad    "
                if ($sub.Count) {
                    $lines += "${pad}With .$name"
                    $lines += $sub
                    $lines += "${pad}End With"
                }
            }
        } elseif ($prof.Kind -eq 'ucsPrkCollection') {
            #--- collections are emitted by the caller, which knows the Add
            #--- signature for the item class
        } elseif ($prof.Kind -in @('ucsPrkPicture','ucsPrkObject')) {
            if ($val) { $script:skipped += "$cls.$name ($($prof.Kind))" }
        } elseif ($prof.CanWrite) {
            $lit = Format-Literal $val $prof
            if (-not [string]::IsNullOrWhiteSpace($lit)) { $lines += "$pad.$name = $lit" }
        }
    }
    return $lines
}

$snap = Get-Content $Snapshot -Raw | ConvertFrom-Json
$class = $snap.class
if (-not $ControlName) { $ControlName = $snap.control }
$props = $snap.props
if (-not $props) { throw "snapshot has no props section (raw-only capture): $Snapshot" }

$sb = New-Object Text.StringBuilder
if ($Out) {
    #--- a .bas needs its module name as the first line, before anything else
    $null = $sb.AppendLine("Attribute VB_Name = ""$([IO.Path]::GetFileNameWithoutExtension($Out))""")
    $null = $sb.AppendLine("Option Explicit")
    $null = $sb.AppendLine("DefObj A-Z")
    $null = $sb.AppendLine("")
}
$null = $sb.AppendLine("'--- generated by tools\GenSample.ps1 from $([IO.Path]::GetFileName($Snapshot))")
$null = $sb.AppendLine("'--- design-time state of $ControlName in sample ""$($snap.sample)"", form $($snap.form)")
$null = $sb.AppendLine("'--- regenerate rather than edit by hand")
if (-not $SubName) { $SubName = "Setup$ControlName" }
$null = $sb.AppendLine("Public Sub $SubName($ControlName As $class)")
$null = $sb.AppendLine("    With $ControlName")

foreach ($l in (Emit-Object $class $props "        ")) { $null = $sb.AppendLine($l) }

#--- collections last: columns carry the bulk of a sample's layout
foreach ($name in ($props.PSObject.Properties.Name | Sort-Object)) {
    $prof = $profs[$class][$name]
    if (-not $prof -or $prof.Kind -ne 'ucsPrkCollection') { continue }
    $value = $props.$name
    if (-not $value) { continue }
    $itemClass = $prof.ItemClass
    #--- a collection with persistable properties of its own snapshots as an
    #--- object carrying them plus an "items" array, not as a bare array
    if ($value -is [array]) {
        $items = $value
        $own = $null
    } else {
        $items = $value.items
        $own = $value
    }
    #--- each collection's Add takes its own arguments, mirroring the
    #--- mappings in tools\common\mdImport.bas
    $consumed = @{
        'JSColumn'       = @('Caption','ColumnType','EditType','Key')
        'JSFmtCondition' = @('ColIndex','Operator','Value1','Value2','Key')
        'JSGroup'        = @('ColIndex','SortOrder')
        'JSSortKey'      = @('ColIndex','SortOrder')
        'JSValueItem'    = @('Value','Text','IconIndex')
        'JSFormatStyle'  = @('Name')
    }
    if ($itemClass -eq 'JSGridImage') {
        $script:skipped += "$name (pictures cannot be generated as literals)"
        continue
    }
    $sysStyles = @('Default','OddRow','EvenRow','RowGroup','PreviewRow','SelectedRow')
    if ($own) {
        $lines = Emit-Object $prof.Type $own "            "
        if ($lines.Count) {
            $null = $sb.AppendLine("        With .$name")
            foreach ($l in $lines) { $null = $sb.AppendLine($l) }
            $null = $sb.AppendLine("        End With")
        }
    }
    foreach ($item in $items) {
        $args = @()
        foreach ($a in $consumed[$itemClass]) {
            $ap = $profs[$itemClass][$a]
            $lit = if ($ap) { Format-Literal $item.$a $ap } else { $null }
            if ($null -eq $lit) { $lit = $(if ($a -in @('Caption','Key','Text','Name')) { '""' } else { '0' }) }
            $args += $lit
        }
        if ($itemClass -eq 'JSFormatStyle' -and $item.Name -in $sysStyles) {
            #--- system styles survive Clear and cannot be re-added
            $add = ".$name.Item(""$($item.Name)"")"
        } elseif ($consumed.ContainsKey($itemClass)) {
            $add = ".$name.Add(" + ($args -join ', ') + ")"
        } else {
            $add = ".$name.Add()"
        }
        $null = $sb.AppendLine("        With $add")
        foreach ($l in (Emit-Object $itemClass $item "            " $consumed[$itemClass])) { $null = $sb.AppendLine($l) }
        $null = $sb.AppendLine("        End With")
    }
}

$null = $sb.AppendLine("    End With")
$null = $sb.AppendLine("End Sub")

if ($script:skipped.Count) {
    Write-Host "skipped (needs hand-written code): $($script:skipped -join ', ')" -ForegroundColor DarkYellow
}
if ($Out) {
    [IO.File]::WriteAllText($Out, $sb.ToString(), [Text.Encoding]::Default)
    Write-Host "wrote $Out"
} else {
    $sb.ToString()
}
