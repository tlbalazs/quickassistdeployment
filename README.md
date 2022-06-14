# quickassistdeployment

I inspired by DroidKid's script when i wrote this. Original script located here: https://droidkid.net/2022/05/17/using-intune-to-deploy-the-new-microsoft-quick-assist/

1. I downloaded the Quick Assist application bundle installer from the organization's Private Store (Microsoft Store for Business). 
2. I placed this installer on a shared location in a server.
3. The script located on a share as well.
4. I created a Group Policy for the script: Computer Configuration -> Policies -> Windows Settings -> Scripts -> Startup -> PowerShell Scripts
   Note: Dont use FQDN when running PowerShell on Startup script because it is broke the running by PowerShell Restriction Policies. Instead of FQDN use the share's full UNC path.

Functions in Script
function MakeLogFolder() 
Create a folder at every day for the new logs and the folder name is the current date.

function LogOutput($Message) 
Create a txt log file for every PC and logging the installation methods. The filename contains the computername. 

function InstallLogOutput($Message) 
Logging all computer install states to one txt file. Contains date, computername, install state, installed version. Works great with Excel to filter data.

function RemoveOldLogs()
Removes the logs that older than 90 days and calls the RemoveLog() function to write its successness to a txt file.

function RemoveLog() 
Logging the success and failed remove attempts in a txt file.
