$lte_file_names = Get-ChildItem -Path "$PSScriptRoot\*" -Include *.lte -Name
$gdt_file_names = Get-ChildItem -Path "$PSScriptRoot\*" -Include *.gdt -Name

$timestamp = Get-Date -Format FileDateTimeUniversal
New-Item -Path "$PSScriptRoot\" -Name "Lufus_umbenannt_$timestamp" -ItemType Directory


foreach ($f in $gdt_file_names){
    Copy-Item `
        -Path "$PSScriptRoot\$($f)" `
        -Destination "$PSScriptRoot\Lufus_umbenannt_$timestamp\$(Get-NewGdtFileName($f))"
}

foreach ($f in $lte_file_names){
    Copy-Item `
        -Path "$PSScriptRoot\$($f)" `
        -Destination "$PSScriptRoot\Lufus_umbenannt_$timestamp\$(Get-NewLteFileName($f))"
}

function Get-NewGdtFileName($file){
    $id = ""
    $date = ""
    $time = ""
    $datetime_f = ""

    $path = ([System.String]::Concat($PSScriptRoot, "\", $file))
    Get-Content -Path $path | ForEach-Object {
        if($_ -match '^0163000\w+$'){
            $id = $_.Substring(7)
        }
        if($_ -match '^0\d{2}6228Testdatum\s+.*$'){
            # line begins with "0416228" if only one and "0686228" if multiple tests in report
            # use first match, in case there were multiple tests of the same date on the report
            $date = (($_ -split '\s+') -match '\d\d\.\d\d\.\d\d')[0]
        }
        if($_ -match '^0\d{2}6228Testzeit\s+.*$'){
            # line begins with "0416228" if only one and "0686228" if multiple tests in report
            # use first match, in case there were multiple tests of the same date on the report
            $time = (($_ -split '\s+') -match '\d\d:\d\d')[0]
        }
    }
    
    if ($date -and $time){
        $datetime_f = [datetime]::ParseExact(
            "$date$time",
            'dd.MM.yyHH:mm',
            [Globalization.CultureInfo]::InvariantCulture
        )
    }
    
    return ("LUFU_" + $id + "_" + $datetime_f.ToString("yyyyMMdd_HHmm") + ".gdt")
}

function Get-NewLteFileName($file){
    $id = ""
    $date = ""
    $time = ""
    $datetime_f = ""

    $path = ([System.String]::Concat($PSScriptRoot, "\", $file))
    Get-Content -Path $path | ForEach-Object {
        if($_ -match '^Name:.*Identifikation:\s\w+$'){
            $id = ($_ -split '\s+')[-1]
        }
        if($_ -match '^Datum\s+\d\d\.\d\d\.\d\d\s*$'){
            $date = ($_ -split '\s+') -match '\d\d\.\d\d\.\d\d'
        }
        if($_ -match '^Zeit\s+\d\d:\d\d:\d\d\s*$'){
            $time = ($_ -split '\s+') -match '\d\d:\d\d:\d\d'
        }
    }
    
    if ($date -and $time){
        $datetime_f = [datetime]::ParseExact(
            "$date$time",
            'dd.MM.yyHH:mm:ss',
            [Globalization.CultureInfo]::InvariantCulture
        )
    }
    
    return ("LUFU_" + $id + "_" + $datetime_f.ToString("yyyyMMdd_HHmm") + ".lte")
}