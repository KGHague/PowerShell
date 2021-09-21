[CmdletBinding()]
param ( 
    [Parameter(Position=0)]
    [Int]$DaysToWarn = 14,
    [Parameter(Position=1)]
    [String]$SupportTeam = "NetStandard Support at 913-428-4200",
    [Parameter(Position=2)]
    [String]$From = "Password AutoBot <noreply@netstandard.com>",
    [Parameter(Position=3)]
    [String]$Subject = "FYI - Your account password will expire soon",
    [Parameter(ParameterSetName='AutoSMTP',Mandatory=$True)]
    [Switch]$AutoSMTPServer,
    [Parameter(ParameterSetName='ManualSMTP',Mandatory=$True)]
    [String]$SMTPServer,
    [String]$TestRecipient,
    [Switch]$WhatIf
)

function PreparePasswordPolicyMail ($ComplexityEnabled,$MaxPasswordAge,$MinPasswordAge,$MinPasswordLength,$PasswordHistoryCount)            
{            
    $verbosemailBody = "<p class=MsoNormal>&nbsp;</p><p class=MsoNormal>Below is a summary of the requirements for your new password:</p>`r`n<ul>`r`n"            
    $verbosemailBody += "<li class=MsoNormal>Your password must be changed every <b>" + $MaxPasswordAge + "</b> days.</li>`r`n"            
    If ($ComplexityEnabled) {
        $verbosemailBody += "<li class=MsoNormal>Your new password cannot contain any part of your name or username and must contain 3 of the 4 character types:<ul><li class=MsoNormal>Uppercase letters</li><li class=MsoNormal>Lowercase letters</li><li class=MsoNormal>Numbers</li><li class=MsoNormal>Symbols</li></ul>`r`n"
    }
    If ($MinPasswordLength -gt 0) {
        $verbosemailBody += "<li class=MsoNormal>Your new password must be at least <b>" + $MinPasswordLength + "</b> characters long.</li>`r`n"
    }
    If ($PasswordHistoryCount -gt 0) {
        $verbosemailBody += "<li class=MsoNormal>Your new password cannot be the same as the last <b>" + $PasswordHistoryCount + "</b> passwords that you have used.</li>`r`n"
    }
    If ($MinPasswordAge -eq 1) {
        $verbosemailBody += "<li class=MsoNormal>You must wait <b>" + $MinPasswordAge + "</b> day before you can change your password again.</li>`r`n"
    }
    If ($MinPasswordAge -gt 1) {
        $verbosemailBody += "<li class=MsoNormal>You must wait <b>" + $MinPasswordAge + "</b> days before you can change your password again.</li>`r`n"
    }
    $verbosemailBody += "</ul>`r`n"
    return $verbosemailBody            
}  

$header = '<html>

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:Wingdings;
	panose-1:5 0 0 0 0 0 0 0 0 0;}
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0in;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1
	{page:WordSection1;}
 /* List Definitions */
 ol
	{margin-bottom:0in;}
ul
	{margin-bottom:0in;}
-->
</style>

</head>

<body lang=EN-US style=''word-wrap:break-word''>

<div class=WordSection1>
'

$footer = "</div>

</body>

</html>
"

#Import AD Module
Import-Module ActiveDirectory -Verbose:$false

$domainPolicy = Get-ADDefaultDomainPasswordPolicy            
$passwordexpirydefaultdomainpolicy = $domainPolicy.MaxPasswordAge.Days -ne 0            
            
if($passwordexpirydefaultdomainpolicy)            
{            
    $defaultdomainpolicyMaxPasswordAge = $domainPolicy.MaxPasswordAge.Days            
    if($verbose)            
    {            
        $defaultdomainpolicyverbosemailBody = PreparePasswordPolicyMail $PSOpolicy.ComplexityEnabled $PSOpolicy.MaxPasswordAge.Days $PSOpolicy.MinPasswordAge.Days $PSOpolicy.MinPasswordLength $PSOpolicy.PasswordHistoryCount            
    }            
} 

#Find accounts that are enabled and have expiring passwords
$users = Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0} `
 -Properties "Name", "UserPrincipalName", "msDS-UserPasswordExpiryTimeComputed", "mS-DS-ConsistencyGuid" `
 | Where-Object {$_."ms-DS-ConsistencyGuid" -ne $null} | Select-Object -Property "Name", "UserPrincipalName", "SAMAccountName", `
 @{Name = "PasswordExpiry"; Expression = {[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed").ToLongDateString() + " " + [datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed").ToLongTimeString() }}

If ($Users -eq $null) {
    Write-Error "No users found with the selected search criteria."
}

If ($TestRecipient) {
    $Users = $Users | Select -First 1
}

#check password expiration date and send email on match
foreach ($user in $users) {

    $DaysRemaining = (New-TimeSpan -Start $(Get-Date) -End $user.PasswordExpiry).Days

    if ($DaysRemaining -le $DaysToWarn) {

        $EmailBody = $header
        $EmailBody += "<p class=MsoNormal>Greetings $($user.Name),</p>`r`n"
        $EmailBody += "<p class=MsoNormal>&nbsp;</p><p class=MsoNormal>This is an automated password expiration warning.&nbsp; Your password will expire in <b>$DaysRemaining</b> days on <b>$($User.PasswordExpiry)</b>.</p>`r`n"

        $PSO= Get-ADUserResultantPasswordPolicy -Identity $user.SAMAccountName            
        if ($PSO -ne $null) {
            $EmailBody += PreparePasswordPolicyMail $PSO.ComplexityEnabled $PSO.MaxPasswordAge.Days $PSO.MinPasswordAge.Days $PSO.MinPasswordLength $PSO.PasswordHistoryCount            
        }
        else {
            $EmailBody += PreparePasswordPolicyMail $domainPolicy.ComplexityEnabled $domainPolicy.MaxPasswordAge.Days $domainPolicy.MinPasswordAge.Days $domainPolicy.MinPasswordLength $domainPolicy.PasswordHistoryCount            
        }
        
        $EmailBody += "<p class=MsoNormal>&nbsp;</p><p class=MsoNormal>To change your password, press the Ctrl+Alt+Delete keys on your keyboard and select ""Change a password"".</p>`r`n"
        $EmailBody += "<p class=MsoNormal><b>Note:</b> If you are not in the office you must first connect to VPN before changing your password.</p>`r`n"
        $EmailBody += "<p class=MsoNormal>&nbsp;</p><p class=MsoNormal>Please contact $SupportTeam if you need assistance changing your password.</p>`r`n"
        $EmailBody += "<p class=MsoNormal>&nbsp;</p><p class=MsoNormal>DO NOT REPLY TO THIS EMAIL. This is an unattended mailbox.</p>`r`n"
        $EmailBody += $footer
         
        If ($AutoSMTPServer) {
            $SMTPServer = (Resolve-DnsName -Type MX -Name $user.UserPrincipalName.Split("@")[1] | sort Preference)[0].NameExchange
        }

        If ($TestRecipient) {
            $Recipient = "$($user.Name) <$TestRecipient>"
        }
        else {
            $Recipient = "$($user.Name) <$user.UserPrincipalName>"
        }
 
        If (-not $WhatIf.IsPresent) {
            Send-MailMessage -To $Recipient -From $From -SmtpServer $SMTPServer -Subject $Subject -BodyAsHtml $EmailBody
        }
        else {
            Write-Host -ForegroundColor Yellow "WhatIf: Server: $SMTPServer  From: $From  To: $Recipient  Days: $DaysRemaining"
        }
        Write-Verbose "Server: $SMTPServer`r`nFrom: $From`r`nTo: $Recipient`r`nSubject: $Subject`r`nBody:`r`n$EmailBody"
    }
}
