#Variables
$FilterWords = @('WkRls', 'Supp', 'supplemental')
$SubFolders  = @("Bond Investigations", "Bond Supervision", "Mug Shots", "Work Education", "Work Release")

#Load the assembly
[void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

#Create a new FolderBrowserDialog object
$FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog

#Set the initial directory of the dialog
$FolderBrowserDialog.RootFolder = 'MyComputer'

#Show the dialog and store the result in a variable
if ($initialDirectory) { $FolderBrowserDialog.SelectedPath = $initialDirectory }
[void] $FolderBrowserDialog.ShowDialog()
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
        '*WkRls'       { $NewFolder = Join-Path -Path $NewFolder -ChildPath 'Work Release' }
        '*WorkRelease' { $NewFolder = Join-Path -Path $NewFolder -ChildPath 'Work Release' }
        Default        { $NewFolder = $NewFolder }
    }

    #Move the file
    Move-Item -Path $File.FullName -Destination $NewFolder
}
