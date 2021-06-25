$user = "NewSignature.Migration@i4-insight.com"
$pass= 'AcesAndEight$$' | ConvertTo-SecureString -AsPlainText -force
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $user, $Pass

Connect-ExchangeOnline -Credential $cred
(Get-mailbox -resultsize Unlimited).count

Get-transportrule