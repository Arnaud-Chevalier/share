###############################################################################################################
#############################################Variables à modifier##############################################
###############################################################################################################

#Variables de sites et de dossier
$SiteURL = "https://yourtenant.sharepoint.com/sites/yoursite"

#Variable finale du dossier à analyser
$FolderSiteRelativeURL = "/yourlibrary/rootfolder/subfolder"

###############################################################################################################
#############################################Début de l'exécution##############################################
###############################################################################################################

#Variables de nom de fichiers et de directory
$Date = Get-Date -Format "dd_MM_yyyy HH.mm.ss"
$CorrectedName = $FolderSiteRelativeURL -replace "/", "  "
$TempDirectory = $env:TEMP 
$CsvName = "$TempDirectory\Tracker - $Date.csv"
$CsvName2 = "$TempDirectory\Count - $Date.csv"

#Fonction de sélection du dossier de destination
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
#Variable du chemin de sauvegarde du dossier
$TrackerPath = Get-Folder

#Connexion à SharePointPNP
Connect-PnPOnline -Url $SiteURL -Credentials (Get-Credential)

#Variables utiles à PnP
$Web = Get-PnPWeb
$Folder = Get-PnPFolder -Url $FolderSiteRelativeURL
 
#Fonction récursive de liste des fichiers
Function Catch-PnPFolder([Microsoft.SharePoint.Client.Folder]$Folder)
{
	### Début du inception mode ###
    #Obtenir l'url relative du dossier
    $FolderSiteRelativeURL = $Folder.ServerRelativeUrl.Replace($web.ServerRelativeUrl,"")
	Write-Host -f Green "Processing on folder : "($FolderSiteRelativeURL.Substring(20))
    #Recherche les fichiers dans le dossier
    $Files = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType File
	$NameOfFiles = $Folder.ServerRelativeUrl
	$NumberOfFiles = $Files.Count
	Write-Host -f Green "Number of files in folder : $NumberOfFiles"
	
	#Compter les fichiers dans chaque dossier
	If ($NumberOfFiles -ge 0)
	{
	$Count = [PSCustomObject]@{
            Folder = $NameOfFiles
            'Number of Files'  = $NumberOfFiles
        }
	$Count | Select @{Name="Directory";Expression={($_.Folder )}}, 'Number of Files', Comments | Export-Csv -Path $CsvName2 -NoTypeInformation -Append -Force -Encoding UTF8
	
	}
	#Export des fichiers vers le csv
    ForEach ($File in $Files)
    {
        #Exporte les données à la suite dans un csv
	    $File  | Select @{Name="Parent Directory";Expression={($_.ServerRelativeUrl -replace "/[^/]*$", "")}}, Name, TimeLastModified, Comments | Export-Csv -Path $CsvName -NoTypeInformation -Append -Force -Encoding UTF8
	}
 
    #Recherche les sous-dossiers dans le dossier
    $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeURL -ItemType Folder
    Foreach($SubFolder in $SubFolders)
    {
        #Exclure les dossiers "Forms" et ceux avec commançant par "_"
        If(($SubFolder.Name -ne "Forms") -and (-Not($SubFolder.Name.StartsWith("_"))))
        {
            #Rappel la fonction pour la rendre récursive (inception mode)
            Catch-PnPFolder -Folder $SubFolder
	### Fin du inception mode ###
        }
    }
} 
#Appel de la fonction de recherche
Catch-PnPFolder -Folder $Folder

#Signal de fin de recherches des fichiers
Write-Host ""
Write-Host ""
Write-Host -f Blue "Processing and extracting datas, please wait... "

#Changement pour le fichier Tracker.csv
Rename-Item "$CsvName" "$CsvName.old"
Import-Csv -Path "$CsvName.old" | sort "Parent Directory"| Export-Csv -Path "$CsvName" -NoTypeInformation -Encoding UTF8
Remove-Item "$CsvName.old"

#Changement pour le fichier Count.csv
Rename-Item "$CsvName2" "$CsvName2.old"
Import-Csv -Path "$CsvName2.old" | sort "Directory" | Export-Csv -Path "$CsvName2" -NoTypeInformation -Encoding UTF8
Remove-Item "$CsvName2.old"

#Conversion finale des csv en xlsx 
Import-Csv "$CsvName" | Export-Excel -Path "$TrackerPath\$CorrectedName - $Date.xlsx" -AutoSize -TableStyle Medium4 -WorkSheetname "Tracker"
Import-Csv "$CsvName2" | Export-Excel -Path "$TrackerPath\$CorrectedName - $Date.xlsx" -AutoSize -TableStyle Medium4 -WorkSheetname "Count"

#Supression des csv restants
Remove-Item "$CsvName"
Remove-Item "$CsvName2"

#Emettre un bip à la fin de l'exécution
[System.Media.SystemSounds]::Beep.Play()

#Appuie sur une touche avant de quitter la console
Write-Host ""
Write-Host ""
Write-Host "End of script execution" 
Read-Host "Press Enter to open the tracker file and exit script..."
Invoke-Item "$TrackerPath\$CorrectedName - $Date.xlsx"
exit