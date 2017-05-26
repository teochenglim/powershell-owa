## Take credential from XML file, uncomment line 2 for the first time
#$cred=GET-CREDENTIAL | EXPORT-CLIXML cred.[your domain and username].xml
#$cred=GET-CREDENTIAL –Credential "staff\clteo" | EXPORT-CLIXML cred.staff.clteo.xml
$cred=IMPORT-CLIXML "cred.staff.clteo.xml"

## State what you want to do with which server
$owa_url = "http://webmail.ntu.edu.sg"
$mail_to = "clteo@ntu.edu.sg"
$mail_subject = "Test email #1"
$mail_msg = "This is an test email. Please don't floor my mailbox"

## Create an IE ComObject and navigate to your exchange OWA webmail server
$ie = New-Object -ComObject InternetExplorer.Application
$ie.visible = $true
$ie.navigate2($owa_url)

## Filled in username and password and click send button
while($ie.busy) { start-sleep -s 1 }

$username = $ie.document.getElementById("username")
$username.value = $cred.username

$password = $ie.document.getElementById("password")
$password.value = $cred.GetNetworkCredential().Password

$link=$ie.Document.getElementsByTagName("input") | where-object {$_.type -eq "submit"}
$link.click()

## Now it is login goto create new email
while($ie.busy) { start-sleep -s 1 }
$new_mail = $ie.document.links | where-object {$_.id -eq "lnkHdrnewmsg"}
$new_mail.click()

## create an new email and click send
while($ie.busy) { start-sleep -s 2 }
$txtto = $ie.document.getElementById("txtto")
$txtto.value = $mail_to
$txtobj = $ie.document.getElementById("txtsbj")
$txtobj.value = $mail_subject
$msg=$ie.Document.getElementsByTagName("textarea") | where-object {$_.title -eq "Message Body"}
$msg.value = $mail_msg

$sendmail = $ie.document.links | where-object {$_.id -eq "lnkHdrsend"}
$sendmail.click()

# check email
while($ie.busy) { start-sleep -s 2 }
$ie.navigate2($owa_url)
while($ie.busy) { start-sleep -s 2 }

## uncomment below if you want quit the browser
#$ie.quit()
