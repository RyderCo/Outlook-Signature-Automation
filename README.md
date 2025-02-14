# Outlook-Signature-Automation
Powershell script to update both local and web Outlook signatures based on AD information.

Prerequisites:
1. Needs to be run from a computer with Domain Admin access, preferably the Domain Controller
2. An App Organization needs to be created in Entra with a cerificate saved to the local store\
   2a. Link to instructions: [App Only Authentication Powershell](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps)
3. Remote Registry needs to be enabled on remote computers\
   3a. Link to instructions: [Enable Remote Registry](https://docs.recastsoftware.com/help/enable-remote-registry)
4. WinRM needs to be enabled on remote computers\
   4a. Link to instructions: [Enable WinRM](https://support.auvik.com/hc/en-us/articles/204424994-How-to-enable-WinRM-with-domain-controller-Group-Policy-for-WMI-monitoring)

Getting Started:

1. Download the files from the Main branch
2. Fill out the signatureconfig.json file
3. Add an optional image to the template folder
4. Fill out the optional UserData.csv
5. Edit the HTML template as needed

The script runs for all users by default, to choose specific users you can use the "FilterUsers" parameter.

Example: .\Outlook_Signature_Automation.ps1 -FilterUsers "testuser,testuser2@yourcompany.com,testuser3"

To only output the user info, use the "TestUsers" switch to output the list of users to a CSV in the same folder as the script.

Example: .\Outlook_Signature_Automation.ps1 -TestUsers

To test an html signature, use the "TestSignature" switch to output the local htm signature and the web html signature to files in the same folder as the script. 
Note: This will ignore other parameters and will not update any users

Example: .\Outlook_Signature_Automation.ps1 -TestSignature
