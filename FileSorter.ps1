#Variables
$FilterWords = @('WkRls', 'supplemental')
$SubFolders  = @("Bond Investigations", "Bond Supervision", "Interview Sheets", "Mug Shots")

#Load the assembly
[void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

#Create a new FolderBrowserDialog object
$FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog

#Set the initial directory of the dialog
$FolderBrowserDialog.RootFolder = 'MyComputer'

#Show the dialog and store the result in a variable
if ($initialDirectory) { $FolderBrowserDialog.SelectedPath = $initialDirectory }
[void] $FolderBrowserDialog.ShowDialog()

#If the user didn't select a folder, exit the script
If ($FolderBrowserDialog.SelectedPath -eq '') {
    Write-Host 'No Folder Selected'
    Exit
}
$FolderPath = $FolderBrowserDialog.SelectedPath

#Grab All Files Under the Folder
$Files = Get-ChildItem -Path $FolderPath -Recurse -File

#Loop through each file, if needed create a new folder and move the file
ForEach ($File in $Files) {
    #Get the file base name, without the extension
    $BaseName = $File.BaseName

    #Clean the base name of Numbers
    If ($BaseName -match '^(?<name>.+?)\d+') {
        $BaseName = $Matches.Name
    }

    #Clean the base name of Exclude Words
    ForEach ($ExcludeWord in $FilterWords) {
        $BaseName = $BaseName -replace $ExcludeWord, ''
    }

    #Create a new folder with the cleaned base name
    $NewFolder = Join-Path -Path $FolderPath -ChildPath $BaseName
    If (-Not (Test-Path -Path $NewFolder)) {
        New-Item -Path $NewFolder -ItemType Directory
    }

    #Create subfolders if they don't exist
    ForEach ($SubFolder in $SubFolders) {
        $SubFolderPath = Join-Path -Path $NewFolder -ChildPath $SubFolder
        If (-Not (Test-Path -Path $SubFolderPath)) {
            New-Item -Path $SubFolderPath -ItemType Directory
        }
    }

    #Move the file to the new folder base on file name
	Switch -WildCard ($File.BaseName) {
    	'*Supp*'     	 { $NewFolder = Join-Path -Path $NewFolder -ChildPath 'Bond Investigations' }
    	'*supplemental*' { $NewFolder = Join-Path -Path $NewFolder -ChildPath 'Bond Investigations' }
    	Default      	 { $NewFolder = Join-Path -Path $NewFolder -ChildPath 'Bond Investigations' }
	}

    #Move the file
    Move-Item -Path $File.FullName -Destination $NewFolder
}
