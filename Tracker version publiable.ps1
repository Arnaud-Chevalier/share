###############################################################################################################
#############################################Modify $SiteURL and $FolderSiteRelativeURL #######################
###############################################################################################################

#Your SharePoint site
$SiteURL = "https://yourtenant.sharepoint.com/sites/yoursite"

#Your folder
$FolderSiteRelativeURL = "/yourlibrary/rootfolder/subfolder"

###############################################################################################################
#############################################Start to execute##################################################
###############################################################################################################

#Folders and files Names
$Date = Get-Date -Format "dd_MM_yyyy HH.mm.ss"
$CorrectedName = $FolderSiteRelativeURL -replace "/", "  "
$TempDirectory = $env:TEMP 
$CsvName = "$TempDirectory\Tracker - $Date.csv"
$CsvName2 = "$TempDirectory\Count - $Date.csv"

#Function to select where the excel is stored
Function Get-Folder($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OutputDirectory = New-Object System.Windows.Forms.FolderBrowserDialog
    $OutputDirectory.Description = "Select a destination folder to store the tracker file : "
    $OutputDirectory.rootfolder = "MyComputer"
    if($OutputDirectory.ShowDialog() -eq "OK")
    {
        $y += $OutputDirectory.SelectedPath
    }
    return $y
}
#Result of the function below
$TrackerPath = Get-Folder

#Connect to SharePoint
Connect-PnPOnline -Url $SiteURL -Credentials (Get-Credential)

#Needed for later
$Web = Get-PnPWeb
$Folder = Get-PnPFolder -Url $FolderSiteRelativeURL
 
#Recursive function to get files in folders and count them
Function Catch-PnPFolder([Microsoft.SharePoint.Client.Folder]$Folder)
{
    $FolderSiteRelativeURL = $Folder.ServerRelativeUrl.Replace($web.ServerRelativeUrl,"")
	Write-Host -f Green "Processing on folder : "($FolderSiteRelativeURL)
    #Find files in folder
    $Files = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType File
	$NameOfFiles = $Folder.ServerRelativeUrl
	$NumberOfFiles = $Files.Count
	Write-Host -f Green "Number of files in folder : $NumberOfFiles"
	
	#Count files in the folders and export to csv
	If ($NumberOfFiles -ge 0)
	{
	$Count = [PSCustomObject]@{
            Folder = $NameOfFiles
            'Number of Files'  = $NumberOfFiles
        }
	$Count | Select @{Name="Directory";Expression={($_.Folder )}}, 'Number of Files', Comments | Export-Csv -Path $CsvName2 -NoTypeInformation -Append -Force -Encoding UTF8
	
	}
	#Get name of files in the folder and export to csv
    ForEach ($File in $Files)
    {
	    $File  | Select @{Name="Parent Directory";Expression={($_.ServerRelativeUrl -replace "/[^/]*$", "")}}, Name, TimeLastModified, Comments | Export-Csv -Path $CsvName -NoTypeInformation -Append -Force -Encoding UTF8
	}
 
    #Look for subfolders in the folder
    $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType Folder
    Foreach($SubFolder in $SubFolders)
    {
        #Exclude folders with the name "Forms" and starting by "_"
        If(($SubFolder.Name -ne "Forms") -and (-Not($SubFolder.Name.StartsWith("_"))))
        {
            #Call-back the function on the subfolder to make it recursive
            Catch-PnPFolder -Folder $SubFolder
        }
    }
} 
#Call the Function on the root folder you selected
Catch-PnPFolder -Folder $Folder

#Tell that the function process ended
Write-Host ""
Write-Host ""
Write-Host -f Blue "Processing and extracting datas, please wait... "

#sorting csv "Tracker" by Parent directory
Rename-Item "$CsvName" "$CsvName.old"
Import-Csv -Path "$CsvName.old" | sort "Parent Directory"| Export-Csv -Path "$CsvName" -NoTypeInformation -Encoding UTF8
Remove-Item "$CsvName.old"

#sorting csv "Count" by Directory
Rename-Item "$CsvName2" "$CsvName2.old"
Import-Csv -Path "$CsvName2.old" | sort "Directory" | Export-Csv -Path "$CsvName2" -NoTypeInformation -Encoding UTF8
Remove-Item "$CsvName2.old"

#Convert and merge both csv to an excel
Import-Csv "$CsvName" | Export-Excel -Path "$TrackerPath\$CorrectedName - $Date.xlsx" -AutoSize -TableStyle Medium4 -WorkSheetname "Tracker"
Import-Csv "$CsvName2" | Export-Excel -Path "$TrackerPath\$CorrectedName - $Date.xlsx" -AutoSize -TableStyle Medium4 -WorkSheetname "Count"

#delete both csv
Remove-Item "$CsvName"
Remove-Item "$CsvName2"

#Play a sound to let you know that your excel is ready
[System.Media.SystemSounds]::Beep.Play()

#Hit a key on your keyboard to open the excel and close powershell
Write-Host ""
Write-Host ""
Write-Host "End of script execution" 
Read-Host "Press Enter to open the tracker file and exit script..."
Invoke-Item "$TrackerPath\$CorrectedName - $Date.xlsx"
exit
