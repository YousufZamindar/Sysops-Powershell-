# Define variables (Replace with your actual credentials and updated email settings)
$Username = "XyzUser"
$OldPassword = Read-Host -Prompt "Enter your old password" -AsSecureString
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $OldPassword
 
# Email Settings (Updated based on provided details)
$SMTPServer = "Enter Your Smtp server host"
$SMTPPort = 25
$EmailSender = "Enter your eamil Address"
$EmailRecipient = "Recepient email address"
$EmailSubject = "New Password"
$EmailBody = "Your new password for Domain is:"
 
function Generate-RandomPassword {
    [CmdletBinding()]
    param (
        [int]$Length = 16
    )
 
    # Character ranges for different types of characters
    $lowercase = 97..122 | ForEach-Object { [char]$_ }
    $uppercase = 65..90 | ForEach-Object { [char]$_ }
    $numbers = 48..57 | ForEach-Object { [char]$_ }
    $specialCharacters = 33..47 + 58..64 + 91..96 + 123..126 | ForEach-Object { [char]$_ }
 
    $passwordArray = @()
 
    # Generate at least one of each type of character
    $passwordArray += Get-Random -InputObject $lowercase
    $passwordArray += Get-Random -InputObject $uppercase
    $passwordArray += Get-Random -InputObject $numbers
    $passwordArray += Get-Random -InputObject $specialCharacters
 
    # Fill the remaining length with random characters
    $remainingLength = $Length - 4
    for ($i = 0; $i -lt $remainingLength; $i++) {
        $passwordArray += Get-Random -InputObject ($lowercase + $uppercase + $numbers + $specialCharacters)
    }
 
    # Shuffle the password array
    $shuffledPassword = ($passwordArray | Get-Random -Count $Length) -join ''
 
    return $shuffledPassword
}
 
function Reset-UserPassword {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory=$true)]
        [System.Security.SecureString]$NewPassword
    )
 
    try {
        Set-ADAccountPassword -Identity $Credential.UserName -NewPassword $NewPassword -Reset
        Write-Output "Password reset successful for $($Credential.UserName)"
    }
    catch {
        Write-Error "Failed to reset password: $_"
    }
}
 
function Send-Email {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SMTPServer,
        [Parameter(Mandatory=$true)]
        [int]$SMTPPort,
        [Parameter(Mandatory=$true)]
        [string]$EmailSender,
        [Parameter(Mandatory=$true)]
        [string]$EmailRecipient,
        [Parameter(Mandatory=$true)]
        [string]$EmailSubject,
        [Parameter(Mandatory=$true)]
        [string]$EmailBody
    )
 
    $SMTPParams = @{
        'SmtpServer' = $SMTPServer
        'Port' = $SMTPPort
        'UseSSL' = $false  # Set to false for TLS disabled
    }
 
    $EmailParams = @{
        'From' = $EmailSender
        'To' = $EmailRecipient
        'Subject' = $EmailSubject
        'Body' = $EmailBody
    }
 
    Send-MailMessage @SMTPParams @EmailParams
}
 
# Generate a new random password
$NewPassword = Generate-RandomPassword -Length 16
 
# Call the function to reset the password
Reset-UserPassword -Credential $Credential -NewPassword (ConvertTo-SecureString -String $NewPassword -AsPlainText -Force)
 
# Send the new password via email
$EmailBody += " $NewPassword"
Send-Email -SMTPServer $SMTPServer -SMTPPort $SMTPPort -EmailSender $EmailSender -EmailRecipient $EmailRecipient -EmailSubject $EmailSubject -EmailBody $EmailBody
 
# Update the old password for next reset
$OldPassword = ConvertTo-SecureString -String $NewPassword -AsPlainText -Force