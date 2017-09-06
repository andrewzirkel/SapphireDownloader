#Sapphire Setup
Import-Module ".\SapphireDownloader"
if (-not $?) {
    log "Error" "SapphireDownloader module not loaded"
    mailIt "$Script $ScriptVersion" $global:mailMessage
    exit
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
    log "Error" "Student Data not downloaded"
    mailIt "$Script $ScriptVersion" $global:mailMessage
    exit
}
$csv | Export-Csv -Path "students.csv" -NoTypeInformation