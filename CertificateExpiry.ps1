###########################################################################################################################################################
# Certificate expiry script                                                                                                                               #
# This scripts checks the expiry dates of certificates and writes to a file (which is then used seperately for comparison against known set of dates.)    #
# Version 1                                                                                                                                               #
# Date: 18/10/2019                                                                                                                                        #
# Author: Lucy McGrother                                                                                                                                  #
#                                                                                                                                                         #
# Acknowledgements - the following have contributed to completing this script.                                                                            #
# http://mekalikot.blogspot.com/2014/08/read-and-get-values-from-excel-file.html for the references used in this script                                   #
# https://blogs.technet.microsoft.com/heyscriptingguy/2015/01/21/adding-and-subtracting-dates-with-powershell/                                            #
# https://stackoverflow.com/questions/15277302/getting-rows-from-a-csv-by-date                                                                            #
# https://social.msdn.microsoft.com/Forums/en-US/d1a338ff-f433-4d65-bb47-c16cb764e0f1/convert-excel-xlsx-to-csv-using-powershell                          #
# https://chris.koester.io/index.php/2015/12/08/use-powershell-select-columns-csv-files/                                                                  #
# https://social.technet.microsoft.com/Forums/en-US/536691c5-4561-41d6-bafd-1164fb2f0ee6/importcsv-without-header-line?forum=winserverpowershell          # 
# https://www.nhaustralia.com.au/blog/PowerShell-Basics-Series-Date-and-Time-Manipulations/                                                               #
# https://social.technet.microsoft.com/Forums/ie/en-US/efa0c21e-2fe8-414f-92ed-feed4ac53363/add-text-to-the-start-of-a-line?forum=winserverpowershell     #
# https://social.technet.microsoft.com/Forums/en-US/11071489-8b87-401d-a9ec-2516c6a97112/addcontent-versus-getdate-text-variable-in-powershell?forum=ITCG #
# https://stackoverflow.com/questions/8501835/making-an-and-statement-to-match-more-than-1-value                                                          #
###########################################################################################################################################################


#Clear all variables before we start
Clear-Host

##########################
# Declare variables here #
##########################

#Set File name here - don't add the .xlsx e.g. $file = "SOC_Certificates"
$File = "SOC_Certificate"
#Set the location of the file above e.g. $loc = "C:\users\MCGROTHERL\Desktop\GetToGreen\Certifcates"
$pwd = "C:\users\MCGROTHERL\Desktop\GetToGreen\Certifcates"
#Set the filepath and name of the file used by EM to raise the call. e.g. $RaiseCallFile = "C:\Users\MCGROTHERL\Desktop\GetToGreen\Certifcates\SOC_Monitoring.txt"
$RaiseCallFile = "C:\Users\MCGROTHERL\Desktop\GetToGreen\Certifcates\SOC_Monitoring.txt"
$WorkingCSV = "C:\Users\MCGROTHERL\Desktop\GetToGreen\Certifcates\SOC_Certificate_stripped.csv"
$Expiring = "Certificate Expiry"
$DateCol = "Cert Expiry" #this is the column name that contains the dates in



#########################
# Fixed Variables       #
# Do not change         #
#########################

#Set Cut off date for which certificates you need e.g. $CutOffDate = (Get-Date).AddDays(+90) - this is looking for certs that expiry in the next 90 days
$CutOffDate90 = (Get-Date).AddDays(+90)
#Set 2nd Cut off date so can identify certificates which are going to expire in 30 days - these need to be raised as a P3.
$CutoffDate30 = (Get-Date).AddDays(+30)
#Create a working csv file - this will be removed when the script finishes
$WorkingCSV = "$pwd\" + "stripped.csv"


##########################################
# Don't change anything below this point #
##########################################


#Set the date with a short date: dd/mm/yyyy
$Today = Get-Date
$Today1 = $Today.ToShortDateString()


#Convert Excel file into CSV Format
Function ExcelCSV ($File)
{
    #Setting params to be used for excel import
    $excelFile = "$pwd\" + $File + ".xlsx"
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($excelFile)
    foreach ($ws in $wb.Worksheets)
    {
        $ws.SaveAs("$pwd\" + $File + ".csv", 6)
    }
    $Excel.Quit()
}


ExcelCSV -File "$File"


#Grab the name of the CSV file we've just created
$filename2 =("$pwd\" + "$File" + ".csv")
