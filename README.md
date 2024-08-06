# DeleteActiveDirectoryHomeFolders
Powershell Tool with GUI to bypass the privileges problem and recursively delete User Home folders created by Active Directory in Windows Server Environments

You need to download and run PSExec in order to run this script as SYSTEM user, which is neccessary for being able to delete User Home folders which are owned by the user.
Here's how you can set it up:

1. Download PsExec: You can download PsExec from the Sysinternals website: [PsExec](https://docs.microsoft.com/en-us/sysinternals/downloads/psexec)
2. Run the Script as SYSTEM: Use PsExec to run PowerShell as the SYSTEM account and then execute your script.

Steps:
1. Download PsExec and extract it to a folder (e.g., C:\PsTools).
2. Open Command Prompt as Administrator: Right-click on the Command Prompt icon and select "Run as administrator".
3. Run PowerShell as SYSTEM:
      psexec -i -s powershell.exe
   (This command will open a new PowerShell window running under the SYSTEM account.)
4. In the now opened PowerShell, navigate to the Script Location and run the script:
      .\DeleteADUserHomeFolders.ps1

Usage:
- You can browse through the folders and select the desired root folder (e.g., D:\Home)
- You can select any number of folders by either using CTRL or Shift
- After pressing "Delete selected" the GUI gets locked and all selected folders will be deleted (this might take a while depending on your system and folder sizes)
- After delete job is done, you receive a prompt, that lists the deleted folder
- In case you choose more than 10 folders for deletion, you will be prompted if you want to export the results into a txt file
  Those results include: Path; Folder Name; Folder Size


That's it. Feel free to use and fork, as long as you don't remove our copyright notice.
