# Frederik Sthen Hansen Specialist academy:


#################################################
# Edit within this area:

#Input the location of excel workbook in the code below: Make sure that the input value is within " "
#  example $myExcelPath= "C:\Users\KOM\Documents\FSH Specialist Academy 2022\Programmeringcase2\opgavefiler/GRI_2017_2020 (1)"


$myExcelPath= "C:\Users\KOM\Documents\FSH Specialist Academy 2022\Programmeringcase2\opgavefiler/GRI_2017_2020 (1)"


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
#example: $myOutputPath= "C:\Users\Documents

$myOutputPath= "C:\Users\KOM\Documents\FSH Specialist Academy 2022\Programmeringcase2\Ny mappe (2)\SpecialisternePDFDownloader\Output-Reports\"


#HUSK AT LAVE EN README TIL STUDENTERMEDHJÆLPERE!!!!#


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

function My-Attempt-Download
{ 
param([bool]$isFirstAttempt,[string]$url, [string]$backupUrl, [string]$fileName, [string]$destination)

if($isFirstAttempt -eq $true){$myMethodUrl=$url} else{ if($backupUrl -ne $null){$myMethodUrl=$backupUrl}}





if ($backupUrl -eq ""){$backupUrl="no available description"}



Write-Output -InputObject "now attempting download with the parameters below:";
Write-Output -InputObject "myMethodURL is:"
Write-Output -InputObject $myMethodUrl
Write-Output -InputObject ""
Write-Output -InputObject "backupUrl is:"
Write-Output -InputObject $backupUrl
Write-Output -InputObject ""
Write-Output -InputObject "Destination is"
Write-Output -InputObject $destination
Write-Output -InputObject ""
Write-Output -InputObject "fileName is:"
Write-Output -InputObject $fileName
Write-Output -InputObject ""


if ($myMethodUrl-eq "" -or $myMethodUrl -eq "no available description"){Write-Output -InputObject "No URL found: Download attempt abandoned"; return}

$Job = Start-BitsTransfer -Source $myMethodUrl -Destination $destination -DisplayName $fileName -Description $backupUrl -Asynchronous -SecurityFlags RedirectPolicyDisallow

 # set max download time to 10 seconds and gives it 60 seconds (the minimum allowed by the BitsJob) to succesfully connect
 $Job= Set-BitsTransfer -BitsJob $Job -MaxDownloadTime 10 -RetryInterval 60 -RetryTimeout 60

 

 Write-Output -InputObject "Bitsjob description:"
 Write-Output -InputObject $Job.Description

 if($isFirstAttempt -eq $true){$myJobs.Add($job)>$null;}else{$myRetryJobs.Add($job)>$null}

 
}


function My-loop-thorugh-BitsJobs
{ 
param([System.Collections.ArrayList]$list)


foreach ($myLittleItem in $list)
{ 

### this loops for every bit transfer
while (($myLittleItem.JobState -eq "Transferring") -or ($myLittleItem.JobState -eq "Connecting")) `
       {
        $transferMessage= "{0} {1}" -f $Job.JobState,$myName;
        Write-Output -InputObject $transferMessage;
        sleep 3;
        

        if ($myLittleItem.JobState -eq "Connecting")
        {
            $connectionAttempts++;
            $transferMessage="{0} {1} {2}"-f $transferMessage,"Connection attempts:",$connectionAttempts;
            Write-Output -InputObject $transferMessage
            if($connectionAttempts -ge 30)
            {
                Write-Output -inputObject "Download timed out! quitting transfer";

                $description=$myLittleItem.Description
                $myName=$myLittleItem.DisplayName
                $myDestination="{0}{1}{2}" -f $myOutputPath,$myName,".PDF";
               $job2= My-Attempt-Download $false $myUrl $description $myName $myDestination; 
                Remove-BitsTransfer -BitsJob $myLittleItem;
            }
        }

      
       } # Poll for status, sleep for 3 seconds, or perform an action.

    Switch($myLittleItem.JobState)
    {    #succesful download
    "Transferred" {$transferMessage="{0} {1}"-f $myLittleItem.DisplayName, "is now fully downloaded!"
     Write-Output -InputObject $transferMessage
    get-bitsTransfer -JobId $myLittleItem.JobId | Complete-BitsTransfer}

    
    #failed
    "Error" 
    {
     # Write-output -inputObject $transfermesage;  
    $myLittleItem | Format-List ;  
    } # List the errors.

    default 
    {
    "retrying with second link"; 
    # Suspend-BitsTransfer -BitsJob $myItem; 
    $description= $myLittleItem.Description
    Write-Output -inputObject "Now writing description";
    Write-Output -InputObject $description;
    $myName=$myLittleItem.DisplayName

    $2ndJob= My-Attempt-Download $false $myUrl $description $myName $myDestination
    } #  Perform corrective action.

    #HUSK AT LAVE EN README TIL STUDENTERMEDHJÆLPERE!!!!#

    }

}

}

function My-add-to-results
{
param ([string]$paramObject,[string]$paramResult,[string]$paramNotes )
#I may need to do some opening operations first

$looper=1
 while($looper -le 3)
 { #switch($looper){} ##CONTINUE HERE!!!!
 $filler;
 $resultSheet.cells.item($resultNumber,$looper)=$filler

 $resultNumber++;

 $looper++;
 }
}

#######
#end of the functions seciton
########


# Opens excel
$ExcelObj = New-Object -comobject Excel.Application

# Make my results spreadsheet
$ExcelResultsBook=$ExcelObj.workbooks.add()
$resultSheet = $ExcelResultsBook.worksheets.item(1)
$resultSheet.name = "Download Results"


$resultSheet.cells.item(1,1) = 'BR-Number'
$resultSheet.cells.item(1,2) = 'Download-result'
$resultSheet.cells.item(1,3) = 'Notes'

$resultNumber=2
#


#Opens Workbook (change the path in string
$ExcelWorkBook = $ExcelObj.Workbooks.Open($myExcelPath)



$ExcelResultsBook
#Opens sheet
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item($mySheet)


$rowsToLoopThrough=( $ExcelWorkSheet.UsedRange.Rows.Count )
# Write-Output -InputObject $rowsToLoopThrough

$myJobs = New-Object System.Collections.ArrayList
$myRetryJobs = New-Object System.Collections.ArrayList 

############
#PLACEHOLDER CODE FOR TESTING!!!

$rowsToLoopThrough= 50

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

 if($myUrl -eq "")
 {
 Write-Output -inputObject "first link is nonexistent, trying second link";  
 $myUrl=$ExcelWorkSheet.cells.Item($myLoopIterator, $mySecondUrlColumn).value2;
# Write-Output -InputObject $myUrl;

 }
    $backupUrl=$ExcelWorkSheet.cells.Item($myLoopIterator, $mySecondUrlColumn).value2
    # Write-Output -InputObject $myDestination
        # start downloading PDF
       # $connectionAttempts=0
       # $transferWait=0

       
        
    #$Job = Start-BitsTransfer -Source $myUrl -Destination $myDestination -DisplayName $myName -Description $backupUrl -Asynchronous -SecurityFlags RedirectPolicyDisallow

  #  write-output -InputObject "Writing the parameters for the attempted download:"
  #  Write-output -InputObject $myUrl
  #  Write-output -InputObject $backupUrl
  #  Write-output -InputObject $myName
  #  Write-output -InputObject $myDestination





 #This  order callsmy attempt download method. the order of the paramaters is Paramount!!
    My-Attempt-Download $true $myUrl $backupUrl $myName $myDestination

    $myLoopIterator++
}

Write-Output -InputObject "Now Looping through myJobs"
My-loop-thorugh-BitsJobs $myJobs;

Write-Output -InputObject "Now Looping through myRetryJobs"
My-loop-thorugh-BitsJobs $myRetryJobs;


#foreach ($myItem in $myJobs)

Write-Output -InputObject "Cleaning away all .tmp files"

#Finds all .tmp files and removes them Link that explains this: https://devblogs.microsoft.com/scripting/how-can-i-use-windows-powershell-to-delete-all-the-tmp-files-on-a-drive/


$myFolderPath=$myOutputPath.Remove($myOutputPath.Length-1, 1)
get-childitem -path $myFolderPath -include *.tmp -Force -Recurse| foreach ($_) {remove-item $_.fullname -Force}
Write-Output -InputObject "Cleaning done"


$doneMessage= "All documents attempted downloaded";
Write-Output -inputObject $doneMessage;




# $myPDFs= New-Object System.Collections.ArrayList

$myPDFs= Get-ChildItem -File -Exclude .tmp -Include *.PDF -Path $myFolderPath -Recurse

foreach ($pdf in $myPDFs)
{
    $myPdfPath="{0}{1}{2}" -f $myOutputPath,$pdf.BaseName,$pdf.Extension ;
    $myPdfContent=Get-Content -Path $myPdfPath -TotalCount 3 ;
    if($myPdfContent -like "*%PDF-*"){$verifiedString= "{0} {1}" -f $pdf.BaseName, "Verified!"; }
    elseif($myPdfContent -like "*<!DOCTYPE html>*")
    {
    $verifiedString="{0} {1}" -f $pdf.BaseName, "is a HTML-file and will be removed";
    
    #Remember to Update my results sheet here!!
     
    Remove-Item $pdf.fullname;
    }
    write-output -InputObject $verifiedString;
    
}

 $doneMessage= "All PDF's verified. The script is now finished!";
Write-Output -inputObject $doneMessage;


