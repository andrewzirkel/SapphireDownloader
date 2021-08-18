#Sapphire Downloader
#1.3

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

#log out to make sure we are in the right building
SDLogout
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
[CmdletBinding()]
param(
[string]$AdditionalFields=$null
)
  login | Out-Null
  #check we are in the right building
  if ($SDCurrentSchool -ne $SDschoold_id){
    SDLogout | Out-Null
    login | Out-Null
  }

$formfields = ""
$formfields+="REPORT_CATEGORY_ID=1&"
$formfields+="REPORT_CODE=DEMO_CUST_LIST&"
$formfields+="SCHOOL_ID=&"
$formfields+="STATUS_FLG=E&"
$formfields+="GRADE_LEVEL=&"
$formfields+="FORMAT=CSV&"
$formfields+="CRLF=Perl&"
$STDENRRPTCOL = "STUDENT_ID,FIRST_NAME,MIDDLE_NAME,LAST_NAME,ADDRESS_1,ADDRESS_CITY,ADDRESS_STATE,ADDRESS_ZIP,PHONE_NO,SSN,GENDER,HOME_ROOM,GRADE_LEVEL,BIRTH_DATE,SCHOOL_ID,ETHNICITY,EMAIL_ADDRESS"
if ($AdditionalFields) {
  $AdditionalFields -split "," | % {
    if ($_ -like "demo_field_id*") {
      $formfields+="CusDemRptCol=$_&"
    }else{
      $STDENRRPTCOL+=",$_"
    }
  }
}
$formfields+="STDENRRPTCOL=" + $STDENRRPTCOL+"&"
# Change the column heading for a column by adding "column_name_of_the_data_you_want":{"DESCRIPTION":"heading_that_you_want"} to the list below.
$formfields+="JSON_STDENRRPTCOL={`"STUDENT_ID`":{`"DESCRIPTION`":`"STUDENT_ID`"},`"HOME_ROOM`":{`"DESCRIPTION`":`"HOME_ROOM`"},`"GRADE_LEVEL`":{`"DESCRIPTION`":`"GRADE_LEVEL`"}}"

$request=$null
$request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/Reports/Reports/DemoCustomListRpt.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
$csv = $request.Content
return $csv
}

function Get-SDClass_Roster {
#for 'CLASS_ROSTER'
[CmdletBinding()]
param(
[string]$CourseID=$null
)

  login | Out-Null
  #check we are in the right building
  if ($SDCurrentSchool -ne $SDschoold_id){
    SDLogout | Out-Null
    login | Out-Null
  }
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
return $csv
}

function get-SDMasterSchedule {
[CmdletBinding()]
param(
[string]$CourseID=$null
)

  login | Out-Null
  #check we are in the right building
  if ($SDCurrentSchool -ne $SDschoold_id){
    SDLogout | Out-Null
    login | Out-Null
  }
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
$formfields += "RptCol=SCHOOL_ID&"
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
return $csv
}

function Get-SDMarkingPeriods {
  login | Out-Null
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/MarkingPeriodInfo.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method Get
  $csv="MPT_MP_CODE,MPT_MP_DESC,mp_start,mp_end,gw_start,gw_end`r`n"
  $currentmp=1
  foreach ($field in $request.InputFields) {
    if ($field.id -and ($field.id -eq "MPT_MP_CODE_$currentmp") -and $field.value){
        $MPCode = $field.value
        [void]$foreach.MoveNext()
        $field = $foreach.Current
        if($field.value) {$MDDesc = $field.value}
        [void]$foreach.MoveNext()
        $field = $foreach.Current
        if($field.value) {$MPStart = $field.value}
        [void]$foreach.MoveNext()
        $field = $foreach.Current
        if($field.value) {$MPEnd = $field.value}
        [void]$foreach.MoveNext()
        $field = $foreach.Current
        if($field.value) {$MPGWStart = $field.value}
        [void]$foreach.MoveNext()
        $field = $foreach.Current
        if($field.value) {$MPGWEnd = $field.value}
        $csv+="$MPCode,$MDDesc,$MPStart,$MPEnd,$MPGWStart,$MPGWEnd`r`n"
        $currentmp++
    }
  }
  return $csv
}

function Get-SDDictionary {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[string]$Dict=$null
)
  #if ([string]::IsNullOrEmpty($Dict)) {#exception }
  login | Out-Null
  #check we are in the right building
  if ($SDCurrentSchool -ne $SDschoold_id){
    SDLogout | Out-Null
    login | Out-Null
  }
  ##we have to hard code the columns :(
  #$dictcolumns = @{}
  #$dictcolumns['Durations']=@("DURATION_CODE","DURATION_DESC","DURATION_GROUP_CODE","STATE_COURSE_SEMESTER_CODE_RID","HIDDEN_FLG","ORDER_NO","ACTIVE_FLG")
  #$dictcolumns['DURATION_MP']=@("DURATION_CODE","MP_CODE","GRADES_FLG","COMMENTS_FLG")
  #$dictcolumns['DURATION_GROUPS']=@("DURATION_GROUP_CODE","DURATION_GROUP_DESC","NUMBER_OF_MARKING_PERIODS","ORDER_NO","ACTIVE_FLG")
  #if(-not $dictcolumns.ContainsKey($Dict)){}#exception}

  #get form columns
  $formfields = ""
  $formfields += "CODEX_CODE=$Dict&"
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/Dictionaries/Main.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
  #use string because of repeated keys in form data
  $formfields = ""
  $formfields += "CODEX_CODE=$Dict&"
  $formfields += "LANGUAGE_CODE=en&"
  $formfields += "PERIOD_PATTERN_CODE=&"
  $request.ParsedHtml.forms['export_form'] | where type -eq "checkbox" | % { $formfields += "columns= " + $_.value + "&" }
  #remove last &
  $formfields = $formfields.Substring(0,$formfields.Length-1)
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/Dictionaries/exportAction.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
  $data = $request.Content
  $enc = [System.Text.Encoding]::ASCII
  $csv = $enc.GetString($data)
  return $csv
}

#returns array of objects with Duration Code, start and end dates
function Get-SDClassDurations {
  login | Out-Null
  #check we are in the right building
  if ($SDCurrentSchool -ne $SDschoold_id){
    SDLogout | Out-Null
    login | Out-Null
  }
  #$data = Get-SDDictionary 'Durations'
  #$DurationsDict = ConvertFrom-Csv $data
  $data = Get-SDDictionary 'DURATION_MP'
  $DurationsMPDict = ConvertFrom-Csv $data
  #$data = Get-SDDictionary 'DURATION_GROUPS'
  #$DurationsGroupsDict = ConvertFrom-Csv $data
  $data = Get-SDMarkingPeriods
  $MPDict = ConvertFrom-Csv $data

  $ClassDurations = @()

  foreach ($record in $DurationsMPDict) {
    if( -not ($ClassDurations | Where-Object -Property 'Duration' -EQ $record.Duration)) {
      $properties = [ordered]@{
                    'Duration'=$record.Duration;
                    'mp_start'='';
                    'mp_end'='';
                    }
      $ClassDurations += New-Object -TypeName psobject -Property $properties
    }
    $MP = $MPDict | Where-Object -Property MPT_MP_CODE -EQ $record.'Marking Period'
    $thisDuraction = $ClassDurations | Where-Object -Property 'Duration' -EQ $record.Duration
    if (-not $thisDuraction.mp_start -or ($(get-date $thisDuraction.mp_start) -gt $(get-date $MP.mp_start))) {$thisDuraction.mp_start = $MP.mp_start}
    if (-not $thisDuraction.mp_end -or ($(get-date $thisDuraction.mp_end) -lt $(get-date $MP.mp_end))) {$thisDuraction.mp_end = $MP.mp_end}
  }
  return $ClassDurations
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

function Get-SDBlackboardConnect {
#for BlackboardConnect
[CmdletBinding()]
param(
#[Parameter(Mandatory=$False)]
[string]$SchoolID
)

login | Out-Null
$formfields = @{}
$formfields['REPORT_CATEGORY_ID']=1
$formfields['REPORT_CODE'] ="EXPORT_BLACKBOARD_CONNECT"
$formfields['SJC_REPORT_ID'] = "EXPORT_BLACKBOARD_CONNECT"
$formfields['Mode'] = "Update"
$formfields['SCHOOL_ID'] = "$schoolID"
#$formfields['STATUS_FLG'] = "E"
#$formfields['GRADE_LEVEL'] = ""
$formfields['FORMAT'] = "CSV"
$formfields['CRLF'] = "Perl"
$request=$null
$request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/Reports/Reports/ExBlackboardConnectRpt.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields -TimeoutSec 300000
$data = $request.Content
$enc = [System.Text.Encoding]::ASCII
$csv = $enc.GetString($data)
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

#returns array of objects with sapphire users
#must have access to Admin > Security, Users and Staff > Users
function Get-SDUsers{
  login | Out-Null
  #capture sid from redirected url https://www.jmcnatt.net/quick-tips/powershell-capturing-a-redirected-url-from-a-web-request/
  #we will get an error about redirection wich can be ignored, set erroraction to silentlycontinue
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/MassUpdateUsers/massUpdateUsers.cfm') -MaximumRedirection 0 -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method Get -ErrorAction SilentlyContinue
  #check if we don't have security rights
  try{$request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/MassUpdateUsers/' + $request.Headers.Location) -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method Get -ErrorAction Stop}
  catch [WebCmdletWebResponseException] {
    write-host "Insufficient permissions to access /Gradebook/CMS/MassUpdateUsers/massUpdateUsers.cfm"
    return
  }
  $request.InputFields | Where-Object {if($_.id -eq 'SID') {$sid=$_.value}}
  $formfields = @{}
  $formfields['FiltersChanged']="true"
  $formfields['BatchSize']=1000
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/MassUpdateUsers/controller/interface.cfm?ACTION=SetOptions' + "&SID=$sid") -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method Post -Body $formfields
  #build user list
  $i=1
  $SDUsers=@()
  while(1) {
    $staffrow=$request.InputFields | Where-Object id -like "*-$i"
    if (-not $staffrow) {break}
    $properties=[ordered]@{}
    foreach ($element in $staffrow) {
      $id=($element.id).replace("id-","")
      $id = $id.replace("-$i","")
      $id = $id.replace("$i","") #mistake in form for staff-r
      $properties[$id]=($element.value)
    }
    $SDUsers+= New-Object -TypeName psobject -Property $properties
    $i++
  }
  return $SDUsers
}

#returns array of objects of additional teachers
#must be in right building
function Get-SDAdditionalTeachers{
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[string]$CourseID,
[Parameter(Mandatory=$True)]
[string]$SectionID
)
  login | Out-Null
  #check we are in the right building
  if ($SDCurrentSchool -ne $SDschoold_id){
    SDLogout | Out-Null
    login | Out-Null
  }
  $additionalTeachers=@()
  $formfields = @{}
  $formfields['COURSE_ID']=$CourseID
  $formfields['SECTION_ID']=$SectionID
  $formfields['Action']='Read'
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/CourseSection.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
  $form=$request.Forms['main_form']
  if($form.Fields['Mode']-ne "Update") {return $false}
  $form=$request.Forms['AdtlTeacher_Form']
  $i=1
  while($form.Fields.ContainsKey("OLD_STAFF_RID_" + $i)) {
    $properties= [ordered]@{
                 'STAFF_RID'=$form.Fields['OLD_STAFF_RID_' + $i];
                 'SHOW_ON_SCHEDULE'=$form.Fields['OLD_SHOW_ON_SCHEDULE_' + $i];
                 'GRADEBOOK_ACCESS'=$form.Fields['OLD_GRADEBOOK_ACCESS_' + $i];
                 'NOTES'=$form.Fields['OLD_COURSE_SECTION_ADTL_TEACHERS_NOTES_' + $i];
                 }
    $i++
    $additionalTeachers += New-Object -TypeName psobject -Property $properties
  }
  return $additionalTeachers
}


function Set-SDStudentEmailAddress{
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[string]$StudentID,
[Parameter(Mandatory=$True)]
[string]$EMail
)
  login | Out-Null
  $formfields = @{}
  $formfields['STUDENT_ID']=$StudentID
  $formfields['Action']='Read'
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/StudentDemographics.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
  $form=$request.Forms['main_form']
  if($form.Fields['Mode']-ne "Update") {return $false}
  #construct update
  #use parsedhtml to apply javascript
  $form=$request.ParsedHtml.forms['main_form']
  $formfields = @{}
  foreach ($element in $form) {
    if ($element.disabled -eq $true) {continue}
    if (($element.type -eq "checkbox") -and (!$element.checked)) {continue}
    if ($element.name){
      $formfields[$element.name]=$element.value
    }
  }
  #remove some fields that the web form doesn't send
  $formfields.Remove('AUTO_STUDENT_ID')
  $formfields.Remove('personsrch_string')
  $formfields.Remove('personsrch_string_2')
  $formfields.Remove('personsrch_string_3')
  $formfields['Action']='Save'
  $formfields['EMAIL_ADDRESS']=$EMail
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/StudentDemographicsAction.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
}

function Get-SDStudentFees{
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[string]$StudentID
)
  login | Out-Null
  $formfields = @{}
  $fees = @() #collection of custom objects
  $formfields['STUDENT_ID']=$StudentID
  $formfields['Action']='Read'
  $formfields['filter_fee_code']=''
  $formfields['filter_fee_type']=''
  $formfields['filter_building']='S'
  $formfields['filter_year']=$SDschool_year
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/StudentDemographicsFees.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
  if ($request.ParsedHtml.title -eq "Error Encountered") { 
    write-host "$StudentID not found"
    return $false
  }
  $form=$request.Forms['main_form']
  if($form.Fields['Mode']-ne "Update") {return $false}
  #construct update
  #use parsedhtml to apply javascript
  $form=$request.ParsedHtml.forms['main_form']
  foreach ($element in $form) {
    if ($element.disabled -eq $true) {continue}
    if (($element.type -eq "checkbox") -and (!$element.checked)) {continue}
    if ($element.name){
      $formfields[$element.name]=$element.value
    }
  }
  for ($i=1; $i -le 100; $i++) {
    if ($formfields["TRANS_ID_$i"] -eq 0) {break}
    $properties = [ordered]@{
        'ID'=$formfields["TRANS_ID_$i"];
        'DATE'=$formfields["TRANS_DATE_$i"];
        'Category'=$formfields["TRANS_CATEGORY_$i"];
        'Deposit'=$formfields["TRANS_Deposit_$i"];
        'Refund'=$formfields["TRANS_REFUND_$i"];
        'Fee'=$formfields["TRANS_FEE_$i"];
        'Payment'=$formfields["TRANS_PAYMENT_$i"];
        }
    $fees+=New-Object -TypeName psobject -Property $properties
  }
  return $fees
}

#Find first free transaction ID in Fees
#parameters: array of form fields
#returns: first free transaction ID as an integer
function findFreeTransID($formfields) {
  for ($i=1; $i -le 100; $i++) {
    if ($formfields["TRANS_ID_$i"] -eq 0) {return $i}
  }
  return 0
}

function Add-SDStudentFee{
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[string]$StudentID,
[string]$Date=$null,
[Parameter(Mandatory=$True)]
[string]$Category,
[string]$Deposit=$null,
[string]$Refund=$null,
[string]$Fee=$null,
[string]$Payment=$null,
[string]$Note=$null
)
  login | Out-Null
  $formfields = @{}
  $formfields['STUDENT_ID']=$StudentID
  $formfields['Action']='Read'
  $formfields['filter_fee_code']=''
  $formfields['filter_fee_type']=''
  $formfields['filter_building']='S'
  $formfields['filter_year']=$SDschool_year
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/StudentDemographicsFees.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
  if ($request.ParsedHtml.title -eq "Error Encountered") { 
    write-host "$StudentID not found"
    return $false
  }
  $form=$request.Forms['main_form']
  if($form.Fields['Mode']-ne "Update") {return $false}
  #construct update
  #use parsedhtml to apply javascript
  $form=$request.ParsedHtml.forms['main_form']
  foreach ($element in $form) {
    if ($element.disabled -eq $true) {continue}
    if (($element.type -eq "checkbox") -and (!$element.checked)) {continue}
    if ($element.name){
      $formfields[$element.name]=$element.value
    }
  }
  $transID=findFreeTransID $formfields
  if ($transID -eq 0) {
    write-host "Could not find free transaction id"
    throw [System.Exception]"Could not find free transaction id."
    return $false
  }
  $formfields['Action']='Save'
  #translate Category to code using Student Fee Dictionary
  $StudentFeeDictCSV=Get-SDDictionary -Dict "STUDENT_FEE_CATEGORIES" | ConvertFrom-Csv
  $CategoryCode=$($StudentFeeDictCSV | where Description -eq $Category).code
  if (-not $CategoryCode) {
    if ($StudentFeeDictCSV | where Code -eq $Category) {
      $CategoryCode=$Category
    } else {
      write-host "$Category not valid in STUDENT_FEE_CATEGORIES Dictionary"
      return $false
    }
  }
  $formfields["TRANS_CATEGORY_$transID"] = $CategoryCode
  if($Date -ne $null) { $formfields["TRANS_DATE_$transID"] = $Date } else { $formfields["TRANS_DATE_$transID"] = get-date -UFormat "%m/%d/%Y" }
  if($Deposit -ne $null) { $formfields["TRANS_DEPOSIT_$transID"] = $Deposit }
  if($Refund -ne $null) { $formfields["TRANS_REFUND_$transID"] = $Refund}
  if($Fee -ne $null) { $formfields["TRANS_FEE_$transID"] = $Fee}
  if($Payment -ne $null) { $formfields["TRANS_PAYMENT_$transID"] = $Payment}
  if($Note -ne $null ) { $formfields["TRANS_DESC_$transID"] = $Note}
  $request = Invoke-WebRequest -Uri ($sapphireURL + '/Gradebook/CMS/StudentDemographicsFeesAction.cfm') -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' -Method POST -Body $formfields
}

function SDLogout {
#need to set tls 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri ($sapphireURL + "/Gradebook/main.cfm?nossl=1&logout=1") -WebSession $script:my_session -UserAgent 'ReportRobot/1.0' | Out-Null
}

function checkLoginParameters {
if ((-Not $SDusername) -or (-Not $SDpassword) -or (-not $sapphireURL) -or (-not $SDdistrict_id) -or (-not $SDschoold_id) -or (-not $SDschool_year)) {return $false}
return $True
}

function login {
#check login parameters are set
if ( -not $(checkLoginParameters) ) {throw [System.Exception]"Login Parameters not set`nMake sure to set Paramenters."}
#need to set tls 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
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
#set school id to check in reports that depend on correct school being logged in.
$script:SDCurrentSchool=$SDschoold_id
}

Export-ModuleMember -Function Set-SDParameters, Get-SDClass_Roster, Get-SDDEMO_CUST_LIST, Get-SDReport, Get-SDConnectEd, Get-SDMasterSchedule, Get-SDMarkingPeriods, Get-SDDictionary, Get-SDClassDurations, Set-SDStudentEmailAddress, Get-SDUsers, Add-SDStudentFee, Get-SDAdditionalTeachers, Get-SDBlackboardConnect, Get-SDStudentFees


