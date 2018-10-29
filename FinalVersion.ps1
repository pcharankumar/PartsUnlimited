#Declare Variaqbles for files
$Date = (Get-Date).ToString("yyyyMMdd-HHmmss")
$FilePath = "D:\GiecoDemo\SonarReports\"
$FileName=$FilePath + "SonarVSTSAnalysisPR" + "_" + $Date + ".xlsx"
$LineSpace = "`r`n"
$TemptxtFile="D:\GiecoDemo\SonarReports\temp.txt"
$TemphtmlFile="D:\GiecoDemo\SonarReports\temp.html"
$DataValidation = ""

#Declare Mail Variables
$Username = "jenkins@primesoft.net";
$password = "GTB&P2nW" | ConvertTo-SecureString -asPlainText -Force
$cred = New-Object -TypeName System.Management.Automation.PSCredential($username,$password)
$subject ="Bugs assigned to you"
$FromEmail = "noreply@nowhere"
$CCEmail = "sbulusu@primesoft.net","pkumar@primesoft.net"
$ValidationTO = "pkumar@primesoft.net"
$ValidationSubject ="Data NOT Found in SonarCube"
$SmtpServer = "smtp.gmail.com"
$portno="587"


#Data 1 - Pull Issues information

       #Data 1 - Pull Issues information

        #$BugURI = "http://192.168.20.184:9000/api/issues/search?resolved=false&types=BUG,CODE_SMELL&fmt=json&ps=500"
        #$BugURI = "http://192.168.20.184:9000/api/issues/search?projects=SonarVSTSAnalysisPR&resolved=false&types=BUG,VULNERABILITY,CODE_SMELL&fmt=json&ps=500"
        $BugURI = "http://192.168.5.7:9000/api/issues/search?projects=SonarVSTSAnalysisPR&resolved=false&types=BUG,VULNERABILITY,CODE_SMELL&fmt=json&ps=500"
        $checkData = Invoke-WebRequest $BugURI | select -ExpandProperty Content | ConvertFrom-Json | Select total
        #$checkData.total
        If ($checkData.total -ge 500)
        {
            $DataValidation ="Data more than 500 records"  +  $LineSpace  + "Here is the URL which used to Pull data : " + $BugURI
            $DataValidation 
        }
        elseif ($checkData.total -eq '')
        {
            $DataValidation ="No Data Found"  +  $LineSpace  + "Here is the URL which used to Pull data : " + $BugURI
            $DataValidation              
        }
        else 
        {
            $BugContent = Invoke-WebRequest $BugURI | select -ExpandProperty content | ConvertFrom-Json
            $Data = $BugContent.issues |   select @{name='Author_new';expression={($_.author)}},
            @{name='Project_New';expression={($_.project)}},
            @{name='Component_New';expression={($_.component).split(':')[-1]}},
            @{name='Type_New';expression={($_.type)}},
            @{name='Message_New';expression={($_.message)}},
            @{name='Sev_New';expression={($_.severity)}}, 
            @{name='Status_New';expression={($_.status)}}, 
            @{name='Assignee_New';expression={($_.assinee)}}, 
            @{name='creationDate_New';expression={($_.creationDate)}}, 
            @{name='updateDate_New';expression={($_.updateDate)}}, 
            @{name='effort_New';expression={($_.effort.replace('min',''))}}, 
            @{name='TD';expression={($_.debt.replace('min',''))}},
            @{name='StartLine';expression={($_.textRange.startline)}},
            @{name='EndLine';expression={($_.textRange.endline)}},
            @{name='LineDiff';expression={($_.textRange.endline) - ($_.textRange.startline)}}
                              
           #$Data | Export-Excel -path $FileName -WorksheetName "ProjectInfo" -IncludePivotTable -PivotDataToColumn -PivotData @{"Type_New"="Count";"TD"="Sum";"MoveToEnd"=$true} -PivotRows "Author_New" -HideSheet "ProjectInfo"
           $Data |Export-Excel -path $FileName -WorksheetName "Issues" -TableName "IssuesData" -TableStyle Light16 -IncludePivotTable  -PivotDataToColumn -PivotData @{"Type_New"="Count";"LineDiff"="Sum";"TD"="Sum"}  -PivotRows "Author_New","Type_New" -PivotFilter "Project_New" -AutoSize
        }

        #Send Mail if No Data Found
        IF ($DataValidation -ne '')
        {
                Send-MailMessage -To $ValidationTO -Subject $ValidationSubject -Body $DataValidation-BodyAsHtml -SmtpServer $SmtpServer -Credential $cred -UseSsl  -From $FromEmail -Port $portno
        }
                
        #Send mail to authors with data
        $AuthorNames = $BugContent.issues | select author -Unique 
        
        #$AuthorNames.GetType()
        foreach ($Au in $AuthorNames)
        {
                $body = $BugContent.issues | sort component | Where-Object {$_.author -eq $Au.author} | Select @{name='component';expression={($_.component).split('/')[-1]} },type,author,creationdate,@{name='message';expression={($_.message).substring(0,50)}},assignee
                #$body = $BugContent.issues | sort component | Where-Object {$_.author -eq $Au.author} | Select type,creationdate,@{name='component';expression={($_.component).split(':')[-1]} },message,assignee
                $body | Format-Table -Wrap  | Out-File $TemptxtFile
        # Convert the fixed width left aligned file into a collection of psobjects
$data = Get-Content $TemptxtFile | Where-Object{![string]::IsNullOrWhiteSpace($_)}

$headerString = $data[0]
$headerElements = $headerString -split "\s+" | Where-Object{$_}
$headerIndexes = $headerElements | ForEach-Object{$headerString.IndexOf($_)}

$results = $data | Select-Object -Skip 2  | ForEach-Object{
    $props = @{}
    $line = $_
    For($indexStep = 0; $indexStep -le $headerIndexes.Count - 2; $indexStep++){
        $value = $null            # Assume a null value 
        $valueLength = $headerIndexes[$indexStep + 1] - $headerIndexes[$indexStep]
        $valueStart = $headerIndexes[$indexStep]
        If(($valueLength -gt 0) -and (($valueStart + $valueLength) -lt $line.Length)){
            $value = ($line.Substring($valueStart,$valueLength)).Trim()
        } ElseIf ($valueStart -lt $line.Length){
            $value = ($line.Substring($valueStart)).Trim()
        }
        $props.($headerElements[$indexStep]) = $value    
    }
    [pscustomobject]$props
} 

# Build the html from the $result
$style = @"
<style>
BODY{font-family:Calibri;font-size:12pt;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse; padding-right:5px}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;color:white;background-color:#FFFFFF }
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:Green}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black}
</style>
"@
$results | Select-Object $headerElements | ConvertTo-Html -Head $style | Set-Content $TemphtmlFile

        $AuthorName = $au.author -replace '\@','%40'

        If ($AuthorName -ne "")
        {
            $URLNEW="http://192.168.20.184:9000/project/issues?id=SonarVSTSAnalysisPR&resolved=false&types=BUG&ps=500&authors=" + $AuthorName
            $bodynew =  "Please find below for the issues assigned to you.`n" + $LineSpace  + (Get-Content $TemphtmlFile -Raw) + $LineSpace + "Click below URL for more details:" + $LineSpace + "`r`n" + $URLNEW
            Send-MailMessage -To $au.author -Subject $subject -Body $bodynew -BodyAsHtml -SmtpServer $SmtpServer -Credential $cred -UseSsl  -From $FromEmail -Cc $CCEmail -Port $portno
        }
        else 
        {
            $URLNEW="http://192.168.20.184:9000/project/issues?id=SonarVSTSAnalysisPR&resolved=false&types=BUG&ps=500"
            $bodynew =  "Please find below for the issues assigned to you.`n" + $LineSpace  + (Get-Content $TemphtmlFile -Raw) + $LineSpace + "Click below URL for more details:" + $LineSpace + "`r`n" + $URLNEW
            Send-MailMessage -To "pkumar@primesoft.net" -Subject $subject -Body $bodynew -BodyAsHtml -SmtpServer $SmtpServer -Credential $cred -UseSsl  -From $FromEmail -Cc $CCEmail -Port $portno        
        }

       }
        
        #Delete temporary files
        Remove-Item -Path $TemphtmlFile
        Remove-Item -Path $TemptxtFile

#Data 2 - Pull Code Coverage Information

        #$MetricURI = "http://192.168.20.184:9000/api/measures/component_tree?component=SonarVSTSAnalysisPR&metricKeys=lines_to_cover,uncovered_lines&ps=500"
        $MetricURI = "http://192.168.5.7:9000/api/measures/component_tree?component=SonarVSTSAnalysisPR&metricKeys=lines_to_cover,uncovered_lines&ps=500"
        $MetricContent = Invoke-WebRequest $MetricURI | select -ExpandProperty content | ConvertFrom-Json

        $MetricContent.components |  ?{$_.measures -notlike '{}' -and $_.path -like '*/*' -and $_.path -ne '/' -and $_.path -notlike 'src/*'} | 
        Select @{name='Project';expression={($_.key).split(':')[1]} },path -ExpandProperty measures | 
        Select Project, Path, Metric, Value | 
        Sort-Object -Property @{Expression = "Metric"; Descending = $False},@{Expression = "Value"; Descending = $true} | 
        ft Project,Path,Value -GroupBy metric

        $a = $MetricContent.components |  ?{$_.measures -notlike '{}' -and $_.path -like '*/*' -and $_.path -ne '/' -and $_.path -notlike 'src/*'} | 
        Select -ExpandProperty measures | 
        ?{$_.Metric -eq 'lines_to_cover'} | 
        Select Project,Path, @{Name ='Lines To Cover'; Expression ={$_.Value -as [int]}}  

        $b = $MetricContent.components |  ?{$_.measures -notlike '{}' -and $_.path -like '*/*' -and $_.path -ne '/' -and $_.path -notlike 'src/*'} | 
        Select -ExpandProperty measures | 
        ?{$_.Metric -eq 'uncovered_lines'} | 
        Select Project,Path, @{Name ='Uncovered Lines'; Expression ={$_.Value -as [int]}} 

        Join-Object -Left $a -Right $b -LeftJoinProperty path -RightJoinProperty path | 
        sort 'Project','Lines To Cover','Uncovered Lines' -Descending |
        Export-XLSX -Path $FileName -WorksheetName "CodeCoverage" -Table -TableStyle Light16 -AutoFit

#Data 3 - Pull Duplications Information

        #$MetricURI = "http://192.168.20.184:9000/api/measures/component_tree?component=SonarVSTSAnalysisPR&branch=master&metricKeys=duplicated_blocks,duplicated_lines&ps=500"
        $MetricURI = "http://192.168.5.7:9000/api/measures/component_tree?component=SonarVSTSAnalysisPR&branch=master&metricKeys=duplicated_blocks,duplicated_lines&ps=500"
        $MetricContent = Invoke-WebRequest $MetricURI |  ConvertFrom-Json
        
        $MetricContent.components |  ?{$_.measures -ne '{}'} | 
        Select @{name='Project';expression={($_.key).split(':')[1]} }, path -ExpandProperty measures | 
        Select Project, Path, Metric, Value | 
        Sort-Object -Property @{Expression = "Metric"; Descending = $False},@{Expression = "Value"; Descending = $true} | 
        ft Path,Value -GroupBy metric
        
        $a = $MetricContent.components |  ?{$_.measures -notlike '{}' -and $_.path -like '*/*' -and $_.path -ne '/' -and $_.path -notlike 'src/*'} | 
        Select -ExpandProperty measures | 
        ?{$_.Metric -eq 'duplicated_lines'} | 
        Select Project,Path, @{Name ='duplicated_lines'; Expression ={$_.Value -as [int]}}  
        
        $b = $MetricContent.components |  ?{$_.measures -ne '{}' -and $_.path -like '*/*' -and $_.path -ne '/' -and $_.path -notlike 'src/*'} | 
        Select -ExpandProperty measures | 
        ?{$_.Metric -eq 'duplicated_blocks'} | 
        Select Project,Path, @{Name ='duplicated_blocks'; Expression ={$_.Value -as [int]}}     
       
        Join-Object -Left $a -Right $b -LeftJoinProperty path -RightJoinProperty path | 
        sort 'project','duplicated_lines','duplicated_blocks' -Descending  |
        Export-XLSX -Path $FileName -WorksheetName "Duplications" -Table -TableStyle Light16 -AutoFit