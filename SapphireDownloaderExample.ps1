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

#Blackboard Connect report
$data = Get-SDBlackboardConnect
#returns 3 group columns so we need to remove header from string and add it back in
$csvdata = $data.Split("`n",2)[1] | ConvertFrom-Csv -header "ReferenceCode","ContactType","FirstName","LastName","Grade","Language","Gender","HomePhone","WorkPhone","MobilePhone","HomePhoneAlt","WorkPhoneAlt","MobilePhoneAlt","SMSPhone","EmailAddress","EmailAddressAlt","Institution","AttendancePhone","AttendancePhoneAlt","Group1","Group2","Group3","HomeAddress","HomeAddress2","HomeCity","HomeState","HomeZip","Terminate"
$csvdata | Export-Csv -Path ".\SDBlackboardConnect.csv" -NoTypeInformation