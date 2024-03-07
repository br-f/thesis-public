cd 'C:\path\to\PCD\directory\'
Get-ChildItem ".\*\LUFU\*.pdf" | Select -exp Name > T:\path\to\output\folder\LUFU_Liste.txt