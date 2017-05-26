# powershell-owa
Use powershell to send email using exchange webmail

## Step 0: clone the repo
c:\code\owa> git clone https://github.com/teochenglim/powershell-owa

## Step 1: Uncomment line 2 for your own to save your password
### Take credential from XML file, uncomment line 2 for the first time

$cred=GET-CREDENTIAL | EXPORT-CLIXML cred.[your domain and username].xml

Example:
$cred=GET-CREDENTIAL –Credential "staff\clteo" | EXPORT-CLIXML cred.staff.clteo.xml

## Step 2: Run your powershell using powershell

PS C:\code\owa> .\webmail1.ps1

## Troubleshooting: For some reason you can't run powershell, make sure it is unrestricted
## You need administrator right to do so

PS C:\code\owa> Set-ExecutionPolicy Unrestricted
