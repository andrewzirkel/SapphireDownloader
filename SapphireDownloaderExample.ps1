#Sapphire Setup
Import-Module ".\SapphireDownloader"
if (-not $?) {
    Write-Host "Error: SapphireDownloader module not loaded"
    Exit
}
$SDusername=""
$SDpassword=""
$sapphireURL=""
$SDdistrict_id=""
$SDschoold_id=""
$SDschool_year=""
Set-SDParameters -Username $SDusername -Password $SDpassword -URL $sapphireURL -DistrictID $SDdistrict_id -SchoolID $SDschoold_id -SchoolYear $SDschool_year

#Custom Demo Report
$csv = $null
#Load students
try { $csv+=Get-SDDEMO_CUST_LIST } catch { write-host $error[0]; exit }
$csv = ConvertFrom-Csv $csv
if(-not $csv){
    Write-Host "Error: Student Data not downloaded"
    exit
}
$csv | Export-Csv -Path "students.csv" -NoTypeInformation