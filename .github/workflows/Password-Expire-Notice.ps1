$DaysToNotify = 5
$SMTPServer = "smtp.state.or.us"
$From = "DoNotReply@no-reply.oregon.gov"
$Subject = "ORA Password Expiration Notice (PAW, PA Program)"

$MaxPwdAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
$ExpiredDate = (Get-Date).addDays(-$MaxPwdAge)

$users = Get-ADUser -Filter {(SamAccountName -Like "ORA.*") -And (Enabled -eq $True) -And (pwdlastset -ne 0)} -Properties PasswordNeverExpires, PasswordLastSet, EmailAddress

Function Email-User ($user,$DaysRemaining) {
    If ($DaysRemaining -lt 0) {
        $DaysRemaining = -$DaysRemaining
        $Body = "Your ORA account expired $DaysRemaining days ago.`nContact The PAW\PA Team @ DAS_DL_OSCIO_ETS_Compute_PAW_PAM@oregon.gov for assistance unlocking your account"
    } elseif ($DaysRemaining -eq 0) {
        $Body = "Your ORA account expires today!`nFor questions, please contact The PAW\PA Team @ DAS_DL_OSCIO_ETS_Compute_PAW_PAM@oregon.gov"
    } else {
        $Body = "Your ORA account expires in $DaysRemaining days.`nFor questions, please contact The PAW\PA Team @ DAS_DL_OSCIO_ETS_Compute_PAW_PAM@oregon.gov"
    }
    Write-Host 
    Write-Host "Email: "$User.EmailAddress
    Write-Host "Subject: $Subject"
    Write-Host "Body: $Body"
    Send-MailMessage -From $From -To $User.EmailAddress -Subject $Subject -Body $Body -SmtpServer $SMTPServer
}

foreach ($user in $users) {
    if ($user.PasswordNeverExpires -eq $False -And $user.EmailAddress) { 
        $DaysRemaining = $user.PasswordLastSet - $ExpiredDate | select -ExpandProperty Days
        if ($DaysRemaining -lt $DaysToNotify) {
            #Write-Host $user.SamAccountName $DaysRemaining
            Email-User $User $DaysRemaining
        }
    }
}