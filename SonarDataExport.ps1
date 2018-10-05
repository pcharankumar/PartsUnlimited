$Date1 = Get-Date -Format ddmmyyyy
$Date2 = Get-Date -Format hhmmss
$FileName="SonarVSTSAnalysisPR" + "_" + $Date1 + "_" + $Date2
$FileName = "D:\GiecoDemo\SonarReports\" + $FileName + ".xlsx"

$uri = "http://192.168.5.7:9000/api/issues/search?id=SonarVSTSAnalysisPR&resolved=false&types=VULNERABILITY,BUG,CODE_SMELL,SECURITY_HOTSPOT&fmt=json"
$content = Invoke-WebRequest $uri | select -ExpandProperty Content | ConvertFrom-Json
$content.issues | sort type,component | select component,type,message,author,creationdate,assignee | 
Export-XLSX -Path $FileName -PivotRows type -PivotValues type -PivotFunction Count 