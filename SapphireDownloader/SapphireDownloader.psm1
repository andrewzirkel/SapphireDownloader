#Sapphire Downloader
#1.1

function Set-SDParameters {
[CmdletBinding()]
param(
[string]$Username,
[string]$Password,
[string]$URL,
[string]$DistrictID,
#[Parameter(Mandatory=$True)]
[string]$SchoolID,
#[Parameter(Mandatory=$True)]
[string]$SchoolYear
)
if($Username) {$script:SDusername=$Username}
if($Password) {$script:SDpassword=$Password}
if($URL) {$script:sapphireURL=$URL}
if($DistrictID) {$script:SDdistrict_id=$DistrictID}
if($SchoolID) {$script:SDschoold_id = $SchoolID}
if($SchoolYear) {$script:SDschool_year = $SchoolYear}
}


#function Get-FormOptions {
##call report options
#$formfields = @{}
#$formfields['REPORT_CATEGORY_ID']=1
#$formfields['REPORT_CODE']='CLASS_ROSTER'
#$request = Invoke-WebRequest -Uri ($sapphireURL + "/Gradebook/CMS/ReportOptions.cfm") -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
#$form = $request.Forms['id_run_report']
##why is it not getting all the fields?
#return $form
#}

function Get-SDDEMO_CUST_LIST {
#for DEMO_CUST_LIST
login | Out-Null
$formfields = @{}
$formfields['REPORT_CATEGORY_ID']=1
$formfields['REPORT_CODE']='DEMO_CUST_LIST'
$formfields['SCHOOL_ID'] = ""
$formfields['STATUS_FLG'] = "E"
$formfields['GRADE_LEVEL'] = ""
$formfields['FORMAT'] = "CSV"
$formfields['CRLF'] = "Perl"
$formfields['STDENRRPTCOL'] = "STUDENT_ID,FIRST_NAME,MIDDLE_NAME,LAST_NAME,ADDRESS_1,ADDRESS_CITY,ADDRESS_STATE,ADDRESS_ZIP,PHONE_NO,SSN,GENDER,HOME_ROOM,GRADE_LEVEL,BIRTH_DATE,SCHOOL_ID"
# Change the column heading for a column by adding "column_name_of_the_data_you_want":{"DESCRIPTION":"heading_that_you_want"} to the list below.
$formfields['JSON_STDENRRPTCOL'] = '{"STUDENT_ID":{"DESCRIPTION":"STUDENT_ID"},"HOME_ROOM":{"DESCRIPTION":"HOME_ROOM"},"GRADE_LEVEL":{"DESCRIPTION":"GRADE_LEVEL"}}'
$request=$null
$request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/Reports/Reports/DemoCustomListRpt.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
$csv = $request.Content
#logout because stuff doesn't work when you switch buildings
SDLogout | Out-Null
return $csv
}

function Get-SDClass_Roster {
#for 'CLASS_ROSTER'
[CmdletBinding()]
param(
[string]$CourseID=$null
)

login | Out-Null
$formfields = @{}
$formfields['REPORT_CATEGORY_ID']=1
$formfields['REPORT_CODE']='CLASS_ROSTER'
$formfields['TEACHER_RID'] = ""
$formfields['COURSE_ID'] = "$CourseID"
$formfields['DEPARTMENT_CODE'] = ""
$formfields['ROOM_CODE'] = ""
$formfields['PERIOD_CODE'] = ""
$formfields['DAY_CODE'] = ""
$formfields['IEP_FLG'] = ""
$formfields['GIEP_FLG'] = ""
#$formfields['DURATIONS'] = "YEAR,FALL,SPR,MP1,MP2,MP3,MP4.SEM1,SEM2,SEM3"
$formfields['SHOW_BLANK_COURSES_FLG'] = "N"
$formfields['SHOW_CURRICULUM_FLG'] = "N"
$formfields['SHOW_SPECIAL_NEEDS_FLG'] = "N"
$formfields['COURSE_ORDER'] = "COURSE_ID"
$formfields['STUDENT_ORDER'] = "NAME"
$formfields['REPORT_FORMAT'] = "CSV"
$formfields['CRLF'] = "Perl"
#$formfields['GRADE_LEVEL => "12,11,10,09",
$formfields['SCHOOL_ID'] = $SDschoold_id
$request=$null
$request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/Reports/Reports/ClassRosterRpt.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
$data = $request.Content
$enc = [System.Text.Encoding]::ASCII
$csv = $enc.GetString($data)
#logout because stuff doesn't work when you switch buildings
SDLogout | Out-Null
return $csv
}

function get-SDMasterSchedule {
[CmdletBinding()]
param(
[string]$CourseID=$null
)

login | Out-Null
#use string because of repeated keys in form data
$formfields = ""
$formfields += "REPORT_CATEGORY_ID=1&"
$formfields += "REPORT_CODE=MASTER_SCHEDULE&"
$formfields += "SJC_REPORT_ID=MASTER_SCHEDULE&"
$formfields += "Mode=Update&"
$formfields += "COURSE_ORDER=COURSE_ID&"
$formfields += "REPORT_FORMAT=CSV&"
$formfields += "CRLF=Perl&"
$formfields += "SCHOOL_ID=$SDschoold_id&"
#display fields
$formfields += "RptCol=COURSE_ID&"
$formfields += "RptCol=SECTION_ID&"
$formfields += "RptCol=COURSE_TITLE&"
$formfields += "RptCol=TIMEPATTERN&"
$formfields += "RptCol=DURATION_CODE&"
$formfields += "RptCol=STAFF_NAME&"
$formfields += "RptCol=STAFF_ID&"
$formfields += "RptCol=ROOM_CODE&"
$formfields += "RptCol=DEPARTMENT_CODE"
$request=$null
$request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/Reports/Reports/MasterScheduleRpt.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
$data = $request.Content
$enc = [System.Text.Encoding]::ASCII
$csv = $enc.GetString($data)
#logout because stuff doesn't work when you switch buildings
SDLogout | Out-Null
return $csv
}

function Get-SDConnectEd {
#for ConnectEd
[CmdletBinding()]
param(
#[Parameter(Mandatory=$True)]
[string]$SchoolID
)

login | Out-Null
$formfields = @{}
$formfields['REPORT_CATEGORY_ID']=1
#$formfields['SCHOOL_ID'] = $SDschoold_id
$formfields['SCHOOL_ID'] = "$schoolid"
$formfields['STATUS_FLG'] = "E"
$formfields['GRADE_LEVEL'] = ""
$formfields['FORMAT'] = "CSV"
$formfields['CRLF'] = "Perl"
$request=$null
$request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/Reports/Reports/ExConnectEdRpt.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields -TimeoutSec 300000
$data = $request.Content
$enc = [System.Text.Encoding]::ASCII
$csv = $enc.GetString($data)
#logout because stuff doesn't work when you switch buildings
SDLogout | Out-Null
return $csv
}

function SDAnalysisLink {
  $request=$null
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/SapphireAnalysisLink.cfm?ContinueAnyway=1') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0'
  $form = $request.Forms['relocate']
  if ($form -eq $null){throw [System.Exception]"Cannot access Analysis"}
  $form.fields.remove('javascript')
  $request = Invoke-WebRequest -Uri ($sapphireURL + $form.Action) -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $form.Fields
  if ($request.ParsedHtml.title -notlike "SapphireAnalysis: Home" ) {throw [System.Exception]"Cannot access Analysis"}
}


function Get-SDReport ($reportID,$format="csv") {
  #for any analysis report represented by ID
  if (-not $reportID ) {throw [System.Exception]"Please include reportID as parameter"}
  login | Out-Null
  SDAnalysisLink #| Out-Null
#  #first generate report
#  $request = Invoke-WebRequest -Uri ($sapphireURL + '/analysis/flow.html?_flowId=viewReportFlow&standAlone=true&ParentFolderUri=/Exports/ConnectEd&reportUnit=/Exports/ConnectEd/ConnectEd_Export') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0'
#  if ($request.StatusCode -ne 200){throw [System.Exception]"Cannot get report"}

  #download report
  $formfields = @{}
  $formfields['_eventId']='export'
  $formfields['_flowExecutionKey']=$reportID
  $formfields['output']=$format
  $formfields['DISTRICT_ID_1']='~NOTHING~'
  $formfields['SCHOOL_YEAR_1']='2013'
  $formfields['SCHOOL_ID_1']='~NOTHING~'
  $request=$null
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/analysis/flow.html/flowFile/ConnectEd_Export.csv') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
  $data = $request.Content
  $enc = [System.Text.Encoding]::ASCII
  $csv = $enc.GetString($data)
  #logout because stuff doesn't work when you switch buildings
  SDLogout | Out-Null
  return $csv
}

function SDLogout {
  Invoke-WebRequest -Uri ($sapphireURL + "/Gradebook/main.cfm?nossl=1&logout=1") -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' | Out-Null
}

function checkLoginParameters {
if ((-Not $SDusername) -or (-Not $SDpassword) -or (-not $sapphireURL) -or (-not $SDdistrict_id) -or (-not $SDschoold_id) -or ($SDschool_year)) {return $false}
return $True
}

function login {
#
#check if we are already logged in
$request = Invoke-WebRequest -Uri ($sapphireURL + "/Gradebook/main.cfm") -WebSession $script:my_session -UserAgent 'ReportRobot/1.0'
if ($request.ParsedHtml.title -like "Select Product - Sapphire Suite" ) {return}
#call login screen
$request = Invoke-WebRequest -Uri ($sapphireURL + "/Gradebook/main.cfm") -SessionVariable script:my_session -UserAgent 'ReportRobot/1.0'
#check for sucsessful login
$form = $request.Forms[0]
$form.fields['j_username'] = $SDusername
$form.fields['j_password'] = $SDpassword
$form.fields['district_id'] = $SDdistrict_id
$form.fields['school_id'] = $SDschoold_id
$form.fields['school_year'] = $SDschool_year
$form.fields.remove('javascript')
$request = Invoke-WebRequest -Uri ($sapphireURL + $form.Action) -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $form.Fields
if ($request.ParsedHtml.title -like "Sapphire Suite - Logon" ) {throw [System.Exception]"Login Failed`nMake sure to set Paramenters."}
}

Export-ModuleMember -Function Set-SDParameters, Get-SDClass_Roster, Get-SDDEMO_CUST_LIST, Get-SDReport, Get-SDConnectEd, get-SDMasterSchedule


