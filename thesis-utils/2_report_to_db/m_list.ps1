cd 'C:\path\to\PCD\directory\'
Get-ChildItem ".\*\N2-MBW\studientauglich\*.pdf" | Select -exp Name > T:\path\to\output\folder\MBW_Liste_studientauglich.txt
Get-ChildItem ".\*\N2-MBW\Studientauglichkeit fraglich\*.pdf" | Select -exp Name > T:\path\to\output\folder\MBW_Liste_fraglich_studientauglich.txt
Get-ChildItem ".\*\N2-MBW\nicht studientauglich\*.pdf" | Select -exp Name > T:\path\to\output\folder\MBW_Liste_nicht_studientauglich.txt