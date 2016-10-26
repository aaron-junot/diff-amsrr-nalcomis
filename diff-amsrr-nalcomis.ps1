$file = 'H:\AMSRR_NMC-PMC_Hi-Pri_Report.xls'
$newname =  $file -replace '\.xls$', '.csv'
$ExcelWB = new-object -comobject excel.application
$Workbook = $ExcelWB.Workbooks.Open($file) 
$Workbook.SaveAs($newname,6)
$Workbook.Close($false)
$ExcelWB.quit()
$file = 'H:\AMSRR_NMC-PMC_Hi-Pri_Report.csv'
$file2 = 'H:\NMCS.txt'$reader = [System.IO.File]::OpenText($file)
$reader2 = [System.IO.File]::OpenText($file2)
$docNums = New-Object System.Collections.ArrayList($null) 
$docNumbers = New-Object System.Collections.ArrayList($null) 
try {
for(;;) {        
$line = $reader.ReadLine() 
$line2 = $reader2.ReadLine()        
if (!($line -eq $null)) {  
if ($line.substring(0,7) -eq "VFA-204") {   
$buno = $line.split(",") -match "16\d\d\d\d"   }   
if($docNumber = $line.split(",") -match "\d\d\d\dG\d\d\d") {   
if ($status = $line.split(",") -match "\d\d\d/\w\w/\w\w\w") {    
$status = $status -replace ‘[/]’,""     
$line = "Document: $docNumber`t BUNO: $buno`t Status: $status"    
$docNums.Add($line + " in AMSRR")   }  }   } 
if (!($line2 -eq $null)) {  
if ($line2.length -gt 2) {
$orgCode = $line2.Substring(0, 3)} 
Else {$orgCode = ""}   if ($orgCode -eq "KA2"){    
if(!($line2.split() -match "162873")){     
if($docNumber2 = $line2.split() -match "\d\d\d\dG\d\d\d") {       
$buno2 = $line2.substring(46,6)      
$status2 = $line2.substring(112,8)      
if($status2 -match "\d\d\dBA\w\w\w" -or $status2 -match "\d\d\dJ\w\w\w") {       
$docNumber2 = -join $docNumber2       
if($docNumber2.length -eq 8){        
$docNumbers.Add("N54076" + $docNumber2 + "xxx")       
}       
elseif($docNumber2.length -eq 9) {        
$docNumbers.Add("N54076" + $docNumber2 + "xx")       
}      }      
if (!($status2 -match "\d\d\dOSSUF")) {      
$line2 = "Document: $docNumber2`t BUNO: $buno2`t Status: $status2"       
$docNums.Add($line2 + " in NALCOMIS")      }     }     }    } } 
if ($line -eq $null -And $line2 -eq $null) {break}    }}
finally {    
$reader.Close() 
$reader2.Close()}
$docNums.Sort()
"*****************************************" > $HOME\AMSRRdifferences.txt
"`n`n`n " >>  $HOME\AMSRRdifferences.txt
for($i=0; $i -lt $docNums.Count; $i+=1) { 
$string1 = $docNums[$i-1] 
$string2 = $docNums[$i+1] 
if (!($string1 -eq $null -or $string2 -eq $null)) {  
if ($docNums[$i].substring(0,50).CompareTo($string2.substring(0,50)) -ne 0 -and $docNums[$i].substring(0,50).CompareTo($string1.substring(0,50)) -ne 0) {   
echo $docNums[$i]`n | Out-File $HOME\AMSRRdifferences.txt -append  } }}
"`n`n`n*****************************************`n`n" >>  $HOME\AMSRRdifferences.txt
"Tracking Numbers`n" >> $HOME\AMSRRdifferences.txt
$url = "https://www.fedex.com/apps/fedextrack/?action=tcn&trackingnumber=n540766082g715xxx&cntry_code=us&shipdate=2016-03-24"
$date = get-date -format u$date = $date.substring(0, 10)
for ($j=0; $j -lt $docNumbers.Count; $j+=1) { 
$trackingURL = $url.substring(0, 65) + $docNumbers[$j] + $url.substring(82, 24) + $date echo $trackingURL`n | Out-File $HOME\AMSRRdifferences.txt -append}
& 'H:\AMSRRdifferences.txt'
