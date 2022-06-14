#Updated on: June 2022
#This script will install the new Quick assist and remove the old version
#Switches used are "-install" and "-uninstall"

#sets switches
param
(
    [Parameter(Mandatory=$false)][Switch]$Install,
    [Parameter(Mandatory=$false)][Switch]$Uninstall,
    [Parameter(ValueFromRemainingArguments=$true)] $args
)

#Global variables
$CurrentDate = Get-Date -Format "yyyy.MM.dd"
$Hostname = "$env:computername"

#Log output results
function LogOutput($Message) {
     $LogFile = "\\SERVERNAME\Distributable_Packages\QuickAssist\Logs\$($CurrentDate)\$($Hostname)_QuickAssist.log"
    "$(get-date -Format 'MM/dd/yyyy HH:mm') $($Message)" | Out-file -FilePath $LogFile -Append -Force
}

function InstallLogOutput($Message) {
    $LogFile = "\\SERVERNAME\Distributable_Packages\QuickAssist\Logs\InstallLog.log"
    "$(get-date -Format 'MM/dd/yyyy HH:mm');$Hostname;$($Message);Version: $($AppXStatus.version)" | Out-file -FilePath $LogFile -Append -Force 
}

function RemoveLog($Message) {
    $LogFile = "\\SERVERNAME\Distributable_Packages\QuickAssist\Logs\RemoveLog.log"
    "$(get-date -Format 'MM/dd/yyyy HH:mm') $($Message)" | Out-file -FilePath $LogFile -Append -Force 
}


function MakeLogFolder() {
     $FolderPath = "\\SERVERNAME\Distributable_Packages\QuickAssist\Logs\$($CurrentDate)"
     If (Test-Path $FolderPath) {
      #Folder exists, do nothing
     }
     Else {
          New-Item -ItemType "Directory" -Path "\\SERVERNAME\Distributable_Packages\QuickAssist\Logs\" -Name "$($CurrentDate)"
     }
}

function RemoveOldLogs() {
    $Counter = 0
    $Folders = Get-ChildItem -Path \\SERVERNAME\Distributable_Packages\QuickAssist\Logs\ -Directory -Name
    ForEach ($Folder in $Folders)
    {
        If ((Get-Date $Folder).DayOfYear+90 -lt (Get-Date $CurrentDate).DayOfYear) {
        Remove-Item -Path "\\SERVERNAME\Distributable_Packages\QuickAssist\Logs\$($Folder)" -Recurse -Force -ErrorAction 'SilentlyContinue'
        $Counter++
        }
    }
   return $Counter
}

#Start of Script
If ($Install){
	Try {
        MakeLogFolder
		LogOutput "***This script is used to install the new Quick Assist app that is from the Microsoft Store. It will also remove the old version***"
		$InstallAppX = Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like "*MicrosoftCorporationII.QuickAssist*"}
		If ($InstallAppX.PackageName -like '*MicrosoftCorporationII.QuickAssist_2022*'){
            Remove-WindowsCapability -Online -Name 'App.Support.QuickAssist~~~~0.0.1.0' -ErrorAction 'SilentlyContinue'
			LogOutput "[Info] Windows Store version of Quick Assist is already installed" 
		}
		Else{
			LogOutput "[Info] Installing the Windows Store version of Quick Assist..."
			Add-AppxProvisionedPackage -online -SkipLicense -PackagePath "\\SERVERNAME\Distributable_Packages\QuickAssist\MicrosoftCorporationII.QuickAssist.AppxBundle"
			LogOutput "[Info] Attempting to remove the old version of Quick Assist..."
			Remove-WindowsCapability -Online -Name 'App.Support.QuickAssist~~~~0.0.1.0' -ErrorAction 'SilentlyContinue'
            LogOutput "[Success] The Windows store version of Quick assist has successfully installed and the old version has been removed." 
		}
		
	} catch [exception] {
		LogOutput "[Error] An error occured installing Quick Assist: $($_.Exception.Message)" 
	}
}

If ($Uninstall){
	Try {
        MakeLogFolder
		LogOutput "***This script is used to uninstall all versions of Microsoft Quick Assist***"
		$AppXStatus = Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like "*MicrosoftCorporationII.QuickAssist*"} 
		#Check to see if the Windows Store version of Quick Assist is installed. Also, lets force an uninstall of the old version just in case 
		If ($AppXStatus.PackageName -notlike '*MicrosoftCorporationII.QuickAssist_2022*'){
			Remove-WindowsCapability -Online -Name 'App.Support.QuickAssist~~~~0.0.1.0' -ErrorAction 'SilentlyContinue'
			LogOutput "[Info] Windows Store version of Quick Assist was not found." 
		}
		#Lets uninstall the Windows Store version of Quick Assist and the old version just in case.
		If ($AppXStatus.PackageName -like '*MicrosoftCorporationII.QuickAssist_2022*'){
            Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like "*MicrosoftCorporationII.QuickAssist_2022.*"} | Remove-AppxProvisionedPackage -Online -AllUsers
			Remove-WindowsCapability -Online -Name 'App.Support.QuickAssist~~~~0.0.1.0' -ErrorAction 'SilentlyContinue'
		}
		LogOutput "[Info] The Windows store version of Quick Assist has successfully been uninstalled." 
	} catch [exception] {
		LogOutput "[Error] An error occurred uninstalling The Windows store version of Quick Assist: $($_.Exception.Message)" 
	}
}

Try {
If (RemoveOldLogs -gt 0){
    RemoveLog "[Success] The logs older than 90 days are successfully removed."
}
} Catch [exception] {
    RemoveLog "[Error] An error occurred when removing the logs older than 90 days. $($_.Exception.Message)"
}

Try {
$AppXStatus = Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like "*MicrosoftCorporationII.QuickAssist*"} -ErrorAction Stop
    If ($AppXStatus.PackageName -like '*MicrosoftCorporationII.QuickAssist_2022*'){
            InstallLogOutput "Quick Assist is installed"
            Exit 0
        } else {
            InstallLogOutput "Quick Assist is uninstalled"
            Exit -1
    }
    } Catch [exception]{
        Write-Error "[Error] $($_.Exception.Message)"
}
