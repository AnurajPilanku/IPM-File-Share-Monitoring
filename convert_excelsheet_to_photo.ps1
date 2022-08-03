#ANURAJ PILANKU
try
{
#$Format=[System.Drawing.Imaging.ImageFormat]::Png
$excel = New-Object -comobject Excel.Application
$FilePath = "\\acdev01\3M_CAC\IPM_FSM\FSM_Excel\IPM_File_Share_Fri_Jan_2022.xlsx"
$YearFolderExist= Test-Path $FilePath
if($YearFolderExist -eq "True")
{

$wb = $excel.Workbooks.Open($FilePath)
$wb.Worksheets("Dashboard").Select()
$wb.ActiveSheet.Range("A1:L21").Select()
$excel.Selection.Copy()
$image=Get-Clipboard -Format Image
$excel.DisplayAlerts=$false
$image.Save("\\acdev01\3M_CAC\IPM_FSM\new.png")#,$Format)
$excel.quit()
Write-Output "success"
}
}
catch
{ 
    Write-Output $_.Exception.Message   
}




