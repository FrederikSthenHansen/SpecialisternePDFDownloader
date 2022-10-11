
Important: 
1.This currently only works in Powershell, when run as Admin, with Developer settings enabled in Windows settings!
2. Its a good idea to enable viewing "hidden" files in the Pathfinder (stifinder). Its a checkbox under "View" ("vis") in the top bar of the Pathfinder.
   This is because the script may leave .tmp files of any failed download attempts 

How the script works:

1. The script uses the values from the "User Input Area" section to open the correct excel worksbook and sheet within that workbook.
2. The script then starts a Bits-transfer Job (download job) for every single row that has been filled in the document. (These take the form of hidden .tmp files)
3. After every Bits-transfer Job has been started it loops through them all starting at the top, 
   to determine if the outcome of the download attempt,and creates PDF's of the succesful downloads.