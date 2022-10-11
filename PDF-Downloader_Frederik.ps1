# Frederik Sthen Hansen Specialist academy:


#################################################
# Edit within this area:

#Input the location of excel workbook in the code below: Make sure that the input value is within " "
#  example $myExcelPath= "C:\Users\KOM\Documents\FSH Specialist Academy 2022\Programmeringcase2\opgavefiler/GRI_2017_2020 (1)"


$myExcelPath= "C:\Users\KOM\Documents\FSH Specialist Academy 2022\Programmeringcase2\opgavefiler/GRI_2017_2020 (1)"


#input name of the sheet within the workbook in the code below. Make sure that the input value is within " "
# example: $mySheet="0"

$mySheet= "0"


#input the path you want the downloaded files deposited at:
#example: $myOutputPath= "C:\Users\KOM\Documents

$myOutputPath= "C:\Users\KOM\Documents\FSH Specialist Academy 2022\Programmeringcase2\Ny mappe (2)\SpecialisternePDFDownloader\Output-Reports\"

#HUSK AT LAVE EN README TIL STUDENTERMEDHJÆLPERE!!!!#

#input start index
$myStartIndex= 

#Input ending index
$myEndIndex=

# This ends the User Input Area
####################################################

#Write-debug -Message $myStartIndex

# if ($myStartIndex -isnot[int]) {Write-Output -InputObject "no start index"}



# Opens excel
$ExcelObj = New-Object -comobject Excel.Application
#Opens Workbook (change the path in string
$ExcelWorkBook = $ExcelObj.Workbooks.Open($myExcelPath)
#Opens sheet
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item($mySheet)


#BR number
#Cells.item(row, Colnmbr) = 
# Write-Output -InputObject $ExcelWorkSheet.cells.Item(4, 1).value2
 

 #col AL
# Write-Output -InputObject $ExcelWorkSheet.cells.Item(4, 38).value2
 #$myUrl=$ExcelWorkSheet.cells.Item(4, 38).value2

 #col AM
 #Write-Output -InputObject $ExcelWorkSheet.cells.Item(4, 39).value2



 #Comment: number of rows is 21058 
 #Write-Output $ExcelWorkSheet.Cells.EntireColumn("A").Count()
 $ExcelWorkSheet.UsedRange.Rows.Count

$rowsToLoopThrough=( $ExcelWorkSheet.UsedRange.Rows.Count )
# Write-Output -InputObject $rowsToLoopThrough


 
$myLoopIterator=1
while($myLoopIterator-le($rowsToLoopThrough))
{
 $myName=$ExcelWorkSheet.cells.Item($myLoopIterator, 1).value2
 $myUrl=$ExcelWorkSheet.cells.Item($myLoopIterator, 38).value2

 $myDestination="{0}{1}{2}" -f $myOutputPath,$myName,".PDF"
# Write-Output -InputObject $myDestination
        # start downloading PDF
        $connectionAttempts=0
        $transferWait=0
 $Job = Start-BitsTransfer -Source $myUrl -Destination $myDestination -DisplayName $myName -Asynchronous 
 
 # set max download time to 10 seconds and gives it 60 seconds (the minimum allowed by the BitsJob) to succesfully connect
 $Job= Set-BitsTransfer -BitsJob $Job -MaxDownloadTime 20 -RetryInterval 120 -RetryTimeout 60

    while (($Job.JobState -eq "Transferring") -or ($Job.JobState -eq "Connecting")) `
       {$transferMessage= "{0} {1}" -f $Job.JobState,$myName;
       Write-Output -InputObject $transferMessage;
       sleep 5;
       if($connectionAttempts>3){$Job.re}
       if ($Job.JobState -eq "Connecting"){$connectionAttempts++}
       } # Poll for status, sleep for 5 seconds, or perform an action.

    Switch($Job.JobState)
    {    #succesful download
    "Transferred" { Write-Output -InputObject $transferMessage
    Complete-BitsTransfer -BitsJob $Job}

    
    #failed
    "Error" {Write-output -inputObject $transfermesage;  $Job | Format-List } # List the errors.
    default {"Other action"} #  Perform corrective action.

    #HUSK AT LAVE EN README TIL STUDENTERMEDHJÆLPERE!!!!#

    }

    $myLoopIterator++
}
$doneMessage= "All documents attempted downloaded"
Write-out -inputObject $doneMessage