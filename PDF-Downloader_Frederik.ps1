# Frederik Sthen Hansen Specialist academy:


Enum myInputType{
Path = 1
File = 2
Number = 3
Text= 4
}
$myPathEnum= [myInputType]::Path
$myFileEnum= [myInputType]::File
$myNumberEnum= [myInputType]::Number
$myTextEnum=[myInputType]::Text


 
function My-set-property
{
param([myInputType]$inputType, [string]$Target)

if ($inputType -eq [myInputType]::File)
{
$userInput= Read-Host -Prompt "Please input the name of the Excel-file to download PDF's from `n"
}
else
{  
  switch([string]$Target)
    {
    "MyExcelPath"
                {$userInput= Read-Host -Prompt "Please input the path to folder containing the excel-file from which to download PDFs and press enter
  `n (Navigate to the file using the pathfinder, go to the bar displaying the sequence of folders and right-click it. `n
   Then select `"copy address`" and paste it below.) `n"}

    "myOutputPath"
                {$userInput= Read-Host -prompt "Please input the path to the folder where you want the PDF's downloaded to and press enter. 
`n((Navigate to the folder using the pathfinder, go to the bar displaying the sequence of folders and right-click it. `n
   Then select `"copy address`" and paste it below.)`n"}

   "myFirstUrlColumn"
{$userInput= read-host -Prompt "Please input the number of the primary column to draw links from `n (column A=1, B=2 etc.) Press enter when done"}

"mySecondUrlColumn"
{$userInput= read-host -Prompt "Please input the number of the secondary column to draw links from `n (column A=1, B=2 etc.) "}

"myNamingColumn"
{$userInput= read-host -Prompt "Please input the number of the column which governs the naming of the downloaded files `n (column A=1, B=2 etc.) Press enter when done "}

"mysheet"
{$userInput= read-host -Prompt "Please input the name of the sheet in the excel-File to draw values from. Press enter when done"}

   


    }
}
My-verify-user-input $userInput $inputType $Target; 
}

function My-verify-user-input
{
param( $ParamInput, [myInputType] $ParamType, [string]$Target )
$verifiedHere=$false;


if ($ParamType -eq [myInputType]::Path -or $ParamType -eq [myInputType]::File)
{ if($paramType -eq[myInputType]::File){ $testPath="{0}{1}{2}"-f $Script:myexcelPath, $userInput,".xlsx";}
else {$testPath=$ParamInput}
$verifiedHere= Test-path -path  $testPath}

if ($ParamType -eq [myInputType]::Number){$ParamInput=[int]$ParamInput; $verifiedHere= $ParamInput -is [int]}

if( $ParamType -eq [myInputType]::Text){$ParamInput=$ParamInput.ToString(); $verifiedHere=$true;}

#
if($verifiedHere-eq $true -and $paramType -eq [myInputType]::Path)
{
switch($target)
    { #is a path switch necesary here?
    "MyExcelPath"
                {$script:myExcelPath="{0}{1}"-f $ParamInput,"\";  }

    "myOutputPath"
                {$script:myOutputPath="{0}{1}"-f $ParamInput,"\"}
    
               # $target="{0}{1}"-f "$",$Target;
               # $myPath="{0}{1}"-f $ParamInput,"\"; 

                #Set-variable -Name $target -Value $myPath -Scope script ;


   }
}
elseif($verifiedHere-eq $true -and $paramType -eq [myInputType]::File){$script:myexcelPath="{0}{1}"-f $Script:myexcelPath,$paramInput}
elseif ($verifiedHere -eq $true)
{switch($target)
    {
    "myFirstUrlColumn"
    {$script:myFirstUrlColumn= $paraminput}

    "mySecondUrlColumn"
    {$script:mySecondUrlColumn= $paraminput}

    "myNamingColumn"
    {$script:myNamingColumn= $paraminput}

    "mysheet"
    {$script:mysheet=$ParamInput}
    }
}



else
{#inform the user of the error and prompt new input.
Write-Output -InputObject "Your input was not valid!"

My-set-property $ParamType $target

}
}

function My-check-downloadStatus
{   Param($ParamJob, [int]$connectionAttempts)
     ### this loops for every bit transfer
    Switch($ParamJob.JobState)
    {    #succesful download
    "Transferred" 
        {
        $transferMessage="{0} {1}"-f $ParamJob.DisplayName, "is now fully downloaded!";
        Write-Output -InputObject $transferMessage;
        get-bitsTransfer -JobId $ParamJob.JobId | Complete-BitsTransfer
        }

        #Download is struggling to connect
     "Connecting"
        {
            $connectionAttempts++;
            $transferMessage="{0} {1}"-f $ParamJob.DisplayName ,"Is struggling to make a connection";
            Write-Output -InputObject $transferMessage

            sleep 60;
            if($connectionAttempts -ge 3)
            {
                Write-Output -inputObject "Download timed out! quitting transfer";

                # make sure this is not a 3rd retry
                If ($ParamJob.Description -cnotlike "*This is a retry!  This is a retry!*")
                {
                    # this is on purpose! 
                    $myUrl= $ParamJob.Description
                    #

                    $description=$ParamJob.Description
                    $myName=$ParamJob.DisplayName
                    $myDestination="{0}{1}{2}" -f $myOutputPath,$myName,".PDF";

                    #Create new download job with 2nd URl
                     My-Attempt-Download $false $myUrl $description $myName $myDestination; 
                     
                    
                }

                #Remove the failed download job from the que
                Remove-BitsTransfer -BitsJob $ParamJob;
                # forced return to avoid recursing
                return;
            }
            #Recurse to check if connection has changed
            My-check-downloadStatus $ParamJob $connectionAttempts;

       }
    
    #failed
    "Error" 
        {
        # Write-output -inputObject $transfermesage;  
         # $ParamJob | Format-List ;  
        } # List the errors.

    default 
        { if($list -ne $myRetryJobs)
            {
            $retryMessage="{0} {1}"-f $ParamJob.Displayname, "retrying with second link"; 
             
            Write-Output -inputObject $retryMessage;
            # Write-Output -InputObject $description;

            $myName=$ParamJob.DisplayName;
            $description=$ParamJob.Description;
            $myDestination="{0}{1}{2}" -f $myOutputPath,$myName,".PDF";
            #############
            #THis Bits transfer needs fixing!!!

            #$true $myUrl $backupUrl $myName $myDestination
            $2ndJob= My-Attempt-Download $false $myUrl $description $myName $myDestination

            #####################
            }
        } #  Perform corrective action.

       # } # Poll for status, sleep for 5 seconds, and recurse.
}

function My-verify-PDFs
{

#Reset my excel row counter to overwrite previous results to reflect any change.
$resultNumber=2

    foreach ($pdf in $myPDFs)
    {
    $verifiedString="";
    $resultString
    $myPdfPath="{0}{1}{2}" -f $myOutputPath,$pdf.BaseName,$pdf.Extension ;
    $myPdfContent=Get-Content -Path $myPdfPath -TotalCount 3 ;
    if($myPdfContent -like "*%PDF-*"){$verifiedString= "{0} {1}" -f $pdf.BaseName, "Verified!"; $resultstring= "Succesful"; }

    elseif($myPdfContent -like "*<!DOCTYPE html>*") {$verifiedString="{0} {1}" -f $pdf.BaseName, "is a HTML-file"; $resultString= "Failed";}

    else{$verifiedString="{0} {1}" -f $pdf.BaseName, "is not a PDF"; $resultString= "Failed";}

    if($resultString -like "*Failed*"){Remove-Item $pdf.fullname;}
    write-output -InputObject $verifiedString;

    # My-add-to-results $pdf.BaseName $resultString $verifiedString;
    
    }
}

function My-Attempt-Download
{ 
    param([bool]$isFirstAttempt,[string]$url, [string]$backupUrl, [string]$fileName, [string]$destination)


    $myMethodUrl=$url;


    if ($backupUrl -like ""){$backupUrl="not available"}
    if ($myMethodUrl -like ""){$myMethodUrl="not available"}

    $myAttemptString= "{0} {1} {2} {3}" -f "now attempting download:",$fileName,"from the url:",$url;
 
    


    if ( $myMethodUrl -like "*not available*")
        {
            $missingUrl="No functional URL found: Download attempt abandoned"
            Write-Output -InputObject $missingUrl;
        
            #assign the backupURL as the URL for the next attempt
            $url=$backupUrl;

            $backupUrl="{0}  {1}" -f $backupUrl, "This is a retry!"
        
            $isFirstAttempt=$false;
        }

    #make sure this is not a 3rd retry
    If ($backupUrl -cnotlike "*This is a retry!  This is a retry!*")
        {   
            Write-Output -InputObject $myAttemptString ;
            Write-Output -InputObject "";
            
            $Job = Start-BitsTransfer -Source $myMethodUrl -Destination $destination -DisplayName $fileName -Description $backupUrl -Asynchronous 

            # set max download time to in seconds and max time to connect in seconds
            $Job= Set-BitsTransfer -BitsJob $Job -MaxDownloadTime 180 -RetryInterval 60 -RetryTimeout 180

            #Add the bitsjob to either the 1st or 2nd attempt collection
            if($isFirstAttempt -eq $true){$myJobs.Add($job)>$null;}
            else{$myRetryJobs.Add($job)>$null}
        }
 
}


function My-loop-thorugh-BitsJobs
{ 
param([System.Collections.ArrayList]$list)


foreach ($myLittleItem in $list)
{ 

    ### this loops for every bit transfer untill it is no longer 
    while (($myLittleItem.JobState -eq "Transferring")) # -or ($myLittleItem.JobState -eq "Connecting")) `
       {
        $transferMessage= "{0} {1}" -f $myLittleItem.JobState,$myLittleItem.DisplayName;
        Write-Output -InputObject $transferMessage;
        
        #
        

        #Wait for 5 seconds
        sleep 5;
        }

      ### this loops for every bit transfer untill it is no longer connecting
     #Check status of the bitsjob
     My-check-downloadStatus $myLittleItem 0
    }

}

}


workflow My-result-writing
{
param($paramPDFs)

$resultBR=$ExcelWorkSheet.cells.Item($resultNumber, $myNamingColumn).value2;
            $myWildBR="{0}{1}{2}"-f "*",$resultBR,"*"; 


            $PdfsString=$myPDFs.Basename ;
            $consoleMessage= "{0}{1}{2}"-f "now checking if ",$resultBR," was succesfully downloaded."
            Write-Host $consoleMessage
            #Check is BR is missing from folder and add negative end result if true: NOt working currently. all PDF are listed as failed.
            if ($paramPDFs.Basename.Contains($resultBR))
                {
                    $myResult= "Successful";$myExplanation="";
                }
            else
                {### Long winded code for insert a new row
                    #add new row to the sheet and record the result
                    $eRow = $resultsheet.cells.item($resultNumber,1).entireRow
                    $active = $eRow.activate()
            
                    $xlShiftDown = [microsoft.office.interop.excel.xlDirection]::xlDown;

                    $active = $eRow.insert($xlShiftDown)
                    ### end of long-winded code for inserting a new row

                    $myResult="Failed";
                    $myExplanation="None of the Urls returned a PDF"; 
                }

            My-add-to-results $resultBR $myResult $myExplanation
            $consoleMessage= "{0}{1}"-f "result found and noted for ",$resultBR
            Write-Host $consoleMessage
}

function My-end-result-writing 
{

    $ExcelResultsBook=$ExcelObj.workbooks.add()
    $resultSheet = $ExcelResultsBook.worksheets.item(1)
    $resultSheet.name ="0" # "Download Results"


    $resultSheet.cells.item(1,1) = 'BR-Number'
    $resultSheet.cells.item(1,2) = 'Download-result'
    $resultSheet.cells.item(1,3) = 'Notes'

    $myResult;
    $myExplanation

    $myPDFs= Get-ChildItem -File -Exclude .tmp -Include *.PDF -Path $myFolderPath -Recurse

    $resultNumber=2
    while($resultNumber -le ($rowsToLoopThrough))
        {
           # $resultBR=$ExcelWorkSheet.cells.Item($resultNumber, $myNamingColumn).value2;
           # $myWildBR="{0}{1}{2}"-f "*",$resultBR,"*"; 


           # $PdfsString=$myPDFs.Basename ;
           # $consoleMessage= "{0}{1}{2}"-f "now checking if ",$resultBR," was succesfully downloaded."
           # Write-Host $consoleMessage
            #Check is BR is missing from folder and add negative end result if true: NOt working currently. all PDF are listed as failed.
           # if ($myPDFs.Basename.Contains($resultBR))
            #    {
            #        $myResult= "Successful";$myExplanation="";
            #    }
            #else
            #    {### Long winded code for insert a new row
            #        #add new row to the sheet and record the result
            #        $eRow = $resultsheet.cells.item($resultNumber,1).entireRow
            #        $active = $eRow.activate()
           # 
            #        $xlShiftDown = [microsoft.office.interop.excel.xlDirection]::xlDown;

             #       $active = $eRow.insert($xlShiftDown)
                    ### end of long-winded code for inserting a new row

              #      $myResult="Failed";
               #     $myExplanation="None of the Urls returned a PDF"; 
            #    }

           # My-add-to-results $resultBR $myResult $myExplanation
        My-result-writing $myPDFs
        $resultNumber++;
        }

     
    $savestring= "{0}{1}" -f $myResultsPath, "results.xlsx";

    #end by saving the file, making it read-only in the process
    $ExcelResultsBook.SaveAs($savestring);
    $ExcelResultsBook.Close
    $excelobj.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelobj)
    # Remove-Variable excelObj
}

function My-add-to-results
{
    param ([string]$paramObject,[string]$paramResult,[string]$paramNotes )
    #I may need to do some opening operations first

$looper=1
 while($looper -le 3)
 { 
 # $filler;

 switch($looper)
 {
 1 {$filler= $paramObject;}
 2{$filler =$paramResult;}
 3{$filler=$paramNotes}
 } 

 
 #Write a row in the excel sheet
 $resultSheet.cells.item($resultNumber,$looper)=$filler

 

 

 $looper++;
 }
 
}

function My-Attempt-Download
{ 
    param([bool]$isFirstAttempt,[string]$url, [string]$backupUrl, [string]$fileName, [string]$destination)
    $myMethodUrl=$url;

    if ($backupUrl -like ""){$backupUrl="not available"}
    if ($myMethodUrl -like ""){$myMethodUrl="not available"}
    
 
    
    if ( $myMethodUrl -like "*not available*")
        {
            $missingUrl="{0} {1}"-f"No functional URL found! swtitching to backup Url for", $fileName;
            Write-Output -InputObject $missingUrl;
        
            #assign the backupURL as the URL for the next attempt
            $myMethodUrl=$backupUrl;

            #Notify the script that this is a retry
            $backupUrl="{0}  {1}" -f $backupUrl, "This is a retry!"
        
            $isFirstAttempt=$false;
        }

    $myAttemptString= "{0} {1} {2} {3}" -f "now attempting download:",$fileName,"from the url:",$myMethodUrl;

    Write-Output -InputObject $myAttemptString ;

    #make sure this is not a 3rd retry
    If ($backupUrl -cnotlike "*This is a retry!  This is a retry!*" -and $myMethodUrl -cnotlike "*not available*")
        {   
            Write-Output -InputObject $myAttemptString ;
            Write-Output -InputObject "";
            
            $Job = Start-BitsTransfer -Source $myMethodUrl -Destination $destination -DisplayName $fileName -Description $backupUrl -Asynchronous 
            # set max download time to in seconds and max time to connect in seconds
            $Job= Set-BitsTransfer -BitsJob $Job -MaxDownloadTime 180 -RetryInterval 60 -RetryTimeout 180
            #Add the bitsjob to either the 1st or 2nd attempt collection

            if($isFirstAttempt -eq $true){$myJobs.Add($job)>$null;}
            else{$myRetryJobs.Add($job)>$null}
        }
    else
    { 
        if($backupUrl -like "*This is a retry!*"){$myAttemptString="{0} {1}{2}"-f"No backup URL found for", $fileName,": All download attempts abandoned" ; }
        else{$myAttemptString="{0} {1}"-f"No URL found for", $fileName}
        Write-Output -InputObject $myAttemptString;
    }
 
}

function My-loop-thorugh-BitsJobs
{ 
param([System.Collections.ArrayList]$list)
foreach ($myLittleItem in $list)
{ 
    ### this loops for every bit transfer untill it is no longer 
    while (($myLittleItem.JobState -eq "Transferring")) # -or ($myLittleItem.JobState -eq "Connecting")) `
        {
        $transferMessage= "{0} {1}" -f $myLittleItem.JobState,$myLittleItem.DisplayName;
        Write-Output -InputObject $transferMessage;

        #Wait for 5 seconds
        sleep 5;
        }
      ### this loops for every bit transfer untill it is no longer connecting
     #Check status of the bitsjob
     My-check-downloadStatus $myLittleItem 0
    }
}

function My-verify-PDFs
{
#Reset my excel row counter to overwrite previous results to reflect any change.
$resultNumber=2
    foreach ($pdf in $myPDFs)
    {
    $verifiedString="";
    $resultString
    $myPdfPath="{0}{1}{2}" -f $myOutputPath,$pdf.BaseName,$pdf.Extension ;
    $myPdfContent=Get-Content -Path $myPdfPath -TotalCount 3 ;
    if($myPdfContent -like "*%PDF-*"){$verifiedString= "{0} {1}" -f $pdf.BaseName, "Verified!"; $resultstring= "Succesful"; }
    elseif($myPdfContent -like "*<!DOCTYPE html>*") {$verifiedString="{0} {1}" -f $pdf.BaseName, "is a HTML-file"; $resultString= "Failed";}
    else{$verifiedString="{0} {1}" -f $pdf.BaseName, "is not a PDF"; $resultString= "Failed";}
    if($resultString -like "*Failed*"){Remove-Item $pdf.fullname;}
    write-output -InputObject $verifiedString;
    # My-add-to-results $pdf.BaseName $resultString $verifiedString;
    
    }
    write-host "Verification done"
}



#######
#end of the functions seciton
########


$script:myExcelPath
$Script:myOutputPath
$Stript:mySheet
[int]$Script:myFirstUrlColumn
[int] $Script:mySecondUrlColumn
[int]$Script:myNamingColumn

#set folder containing excelfile and add the file to the path
 My-set-property $myPathEnum "myExcelPath"; 
 My-set-property $myFileEnum "myExcelPath";

#set ouput folder:
My-set-property $mypathEnum "myOutputPath"; $myResultsPath=$Script:myOutputPath

#set sheet and columns in excel-file 
My-set-property $myTextEnum "mySheet"; My-set-property $myNumberEnum "myFirstUrlColumn"; 
My-set-property $myNumberEnum "mySecondUrlColumn";My-set-property $myNumberEnum "myNamingColumn";





 #

 


# $myFirstUrlColumn=  read-host -prompt "Please input the number of the column to draw links from (column A=1, B=2 etc.) "
 
# $myFirstUrlColumn
# $myResultsPath=$myOutputPath




$timer = [Diagnostics.Stopwatch]::StartNew()


# Opens excel
$ExcelObj = New-Object -comobject Excel.Application

# Make my results spreadsheet

#


#Opens Workbook (change the path in string
$ExcelWorkBook = $ExcelObj.Workbooks.Open($myExcelPath)


#Opens sheet
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item($mySheet)


$rowsToLoopThrough=( $ExcelWorkSheet.UsedRange.Rows.Count )
# Write-Output -InputObject $rowsToLoopThrough

$myJobs = New-Object System.Collections.ArrayList
$myRetryJobs = New-Object System.Collections.ArrayList 

############
#PLACEHOLDER CODE FOR TESTING!!!

# $rowsToLoopThrough= 100

########

$myLoopIterator=2 #start at 2 as 1 is for titles and not values
while($myLoopIterator-le($rowsToLoopThrough))
{
    #reset properties from any previous looping
    $myName="";
    $myUrl="";
    $myDestination="";
    $backupUrl="";

    #Br Number
    $myName=$ExcelWorkSheet.cells.Item($myLoopIterator, $myNamingColumn).value2;
 
    #col AL
    $myUrl=$ExcelWorkSheet.cells.Item($myLoopIterator, $myFirstUrlColumn).value2;

    $myDestination="{0}{1}{2}" -f $myOutputPath,$myName,".PDF";
    # Write-Output -inputObject $myDestination;
    # Write-Output -InputObject $myUrl;

    $backupUrl=$ExcelWorkSheet.cells.Item($myLoopIterator, $mySecondUrlColumn).value2
    # Write-Output -InputObject $myDestination

    # start downloading PDF
    #This  order callsmy attempt download method. the order of the paramaters is Paramount!!
    #  required parametres [bool]$isFirstAttempt,[string]$url, [string]$backupUrl, [string]$fileName, [string]$destination
    My-Attempt-Download $true $myUrl $backupUrl $myName $myDestination
   
    $myLoopIterator++
}

Write-Output -InputObject "Done with initiating all downloads";

Write-Output -InputObject $timer.elapsed.totalseconds;

Write-Output -InputObject "";

Write-Output -InputObject "Now Looping through myJobs"

My-loop-thorugh-BitsJobs $myJobs;

Write-Output -InputObject "Done with Looping through myJobs";

Write-Output -InputObject $timer.elapsed.totalseconds;
Write-Output -InputObject "";

$myFolderPath=$myOutputPath.Remove($myOutputPath.Length-1, 1)

Write-Output -InputObject $timer.elapsed.totalseconds;
Write-Output -InputObject "";

$myPDFs=get-childitem -path $myFolderPath -include *.PDF

My-verify-PDFs;


Write-Output -InputObject "Now Looping through myRetryJobs"
My-loop-thorugh-BitsJobs $myRetryJobs;
Write-Output -InputObject "Done with Looping through myRetryJobs";

Write-Output -InputObject $timer.elapsed.totalseconds;
Write-Output -InputObject "";

$doneMessage= "All documents attempted downloaded";
Write-Output -inputObject $doneMessage;

Write-Output -InputObject $timer.elapsed.totalseconds;
Write-Output -InputObject "";

My-verify-PDFs;

Write-Output -InputObject "Writing download results to excel file"

My-end-result-writing;

Write-Output -InputObject "Cleaning away all .tmp files"

#Finds all .tmp files and removes them Link that explains this: https://devblogs.microsoft.com/scripting/how-can-i-use-windows-powershell-to-delete-all-the-tmp-files-on-a-drive/
get-childitem -path $myFolderPath -include *.tmp -Force -Recurse| foreach ($_) {remove-item $_.fullname -Force}

Write-Output -InputObject "Cleaning done"
 #find en metode til at se .tmp filstørelse og forbinde br-nummer til evaluering i excel arket.

$doneMessage= "All PDF's verified. The script is now finished!";

Write-Output -inputObject $doneMessage;

Write-Output -InputObject $timer.elapsed.totalseconds;

### Code to clean up after testing. DO NOT EXECUTE!
# Get-BitsTransfer -AllUsers| Remove-BitsTransfer
# get-childitem -path $myFolderPath -include * -Force -Recurse| foreach ($_) {remove-item $_.fullname -Force}


###