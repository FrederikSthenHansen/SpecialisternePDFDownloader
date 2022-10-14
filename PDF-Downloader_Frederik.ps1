# Frederik Sthen Hansen Specialist academy:


#################################################
# Edit within this area:

#Input the location of excel workbook in the code below: Make sure that the input value is within " "
#  example $myExcelPath= "C:\Users\KOM\Documents\FSH Specialist Academy 2022\Programmeringcase2\opgavefiler\GRI_2017_2020 (1)"


$myExcelPath= "C:\Users\KOM\Documents\FSH Specialist Academy 2022\Programmeringcase2\opgavefiler\GRI_2017_2020 (1)"


#input name of the sheet within the workbook, that data has to be gathered from, in the code below. Make sure that the input value is within " "
# example: $mySheet="0"

$mySheet= "0"


# Input the number of the column containing the first URL to attempt download from. When counting columns start at 1. for column AL the number is 38
#example:   $myFirstUrlColumn=38
$myFirstUrlColumn=38


# Input the number of the column (vertical) containing the first URL to attempt download from. When counting columns start at 1. for column AL the number is 38
#example:   $mySecondUrlColumn=39
$mySecondUrlColumn=39

#type in the number of the column (vertical) containing the fields by which act as the naming convention for any downloaded PDFs
#example: $myNamingColumn=1
$myNamingColumn=1

#input the path you want the downloaded files deposited at:
#example: $myOutputPath= "C:\Users\Documents"

$myOutputPath= "C:\Users\KOM\Documents\FSH Specialist Academy 2022\Programmeringcase2\Ny mappe (2)\SpecialisternePDFDownloader\Output-Reports\"


#HUSK AT LAVE EN README TIL STUDENTERMEDHJÆLPERE!!!!#

# Here you can input the intended location for the Excel workbook detailing the results from executing the script
# By default the results will be stored in the same folder as the PDF's
# Example of default: $myResultsPath=$myOutputPath
# Example of custom path: $myResultsPath= "C:\Users\Documents\my-results-folder"
$myResultsPath=$myOutputPath

#Use the snippet here in the Powershell console for easy cleaning away of TMP files (remove the # fron each line of the snippet, to activate the code
# and add them again to disable the code (not needed per se)
#Snippet starts below:

# $myFolderPath=$myOutputPath.Remove($myOutputPath.Length-1, 1);
# get-childitem -path $myFolderPath -include *.tmp -Force -Recurse| foreach ($_) {remove-item $_.fullname -Force};

#snippet ends above:
##### 

#input start index
#$myStartIndex= 

#Input ending index
#$myEndIndex=

#comment this out by placing a # in front of it, if you want the script loop all used rows in the excel sheet
#$rowsToLoopThrough=$myEndIndex




# This ends the User Input Area
####################################################




#########
# Functions section:
##########

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
            # Suspend-BitsTransfer -BitsJob $myItem; 
            Write-Output -inputObject $retryMessage;
            # Write-Output -InputObject $description;
            $myName=$ParamJob.DisplayName;
            $description=$ParamJob.Description;
            $myDestination="{0}{1}{2}" -f $myOutputPath,$myName,".PDF";
            #$true $myUrl $backupUrl $myName $myDestination
            $2ndJob= My-Attempt-Download $false $myUrl $description $myName $myDestination
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
            $resultBR=$ExcelWorkSheet.cells.Item($resultNumber, $myNamingColumn).value2;
            $myWildBR="{0}{1}{2}"-f "*",$resultBR,"*"; 


            $PdfsString=$myPDFs.Basename ;

            #Check is BR is missing from folder and add negative end result if true: NOt working currently. all PDF are listed as failed.
            if ($myPDFs.Basename.Contains($resultBR))
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
}



#######
#end of the functions seciton
########

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

 $rowsToLoopThrough= 100

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