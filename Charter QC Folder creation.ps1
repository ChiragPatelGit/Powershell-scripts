Set-ExecutionPolicy -ExecutionPolicy Unrestricted
# Directory to get the workbook file names from

$WorkbookFolderPath = 'C:\Users\patelch\OneDrive - WWT\Charter\200 Site Project\Workbooks' #Path of the excel wookbooks
$ParentFolderPath = 'C:\Users\patelch\OneDrive - WWT\Charter\200 Site Project\QC\inlab\' #path where to create the QC folders

$MyArray = Get-ChildItem -Path $WorkbookFolderPath  -Name *.xls



#Loop that traverses the array of foldernames and then creates the folders
for ($i = 0; $i -lt $MyArray.Length; $i++) {

 $Folderpath = $ParentFolderPath + $MyArray[$i].ToString().trim(".xls")
 
 # If the folder does not exist then create a new folder in the QC inlab folder
 If(!(Test-Path $Folderpath))
 {
    New-Item -Path $Folderpath -ItemType "directory"
  }
 Else
 {
    echo $Folderpath  'already exists'
 }
  
}

$WorkbooksPath = Read-Host -Prompt 'Input the Wookbook Path'

Write-Host "The path of your workbooks is '$WorkbooksPath' "
InputWorkbookPath($WorkbooksPath)

function InputWorkbookPath {
    param ($Path)

    Write-Host "The path of your workbooks is '$Path' "

}
