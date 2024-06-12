param(
    [string]$SamAccountName,
    [int]$PasswordLength = 24,
    [switch]$SplitPassword,
    [switch]$AddEmphasis,
    [switch]$DoNotShow
)

Import-Module "$PSScriptRoot\module.psm1"

$Global:user = Get-ADUser -Identity $SamAccountName -Properties PasswordLastSet
$Global:domain = Get-ADDomain
$password = @(New-Password -Length $PasswordLength)
Set-ADAccountPassword $Global:user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)

if ($SplitPassword.IsPresent) { $password += Split-Password -Password $password }

$i = 0
$layout = Get-Content -Path $PSScriptRoot\layout.html -Encoding UTF8 -Raw
$password | ForEach-Object {
    switch ($i) {
        0 {
            $index = 1
            $Global:pwdIntro = 'The complete password is: '
            $path = "$PSScriptRoot\credentials_full.html"
            $Global:paging = ''
        }
        1 {
            $index = 1
            $Global:pwdIntro = 'The first part of the password is: '
            $path = "$PSScriptRoot\credentials_part1.html"
            $Global:paging = ' [1/2]'
        }
        2 {
            $index = $password[1].Length + 1
            $Global:pwdIntro = 'The last part of the password is: '
            $path = "$PSScriptRoot\credentials_part2.html"
            $Global:paging = ' [2/2]'
        }
    }

    $passwordTable = ConvertTo-HTMLPasswordTable -Password $_ -Index $index
    if ($AddEmphasis.IsPresent) { $passwordTable = Add-Emphasis -HTMLTable $passwordTable }
    $body = $ExecutionContext.InvokeCommand.ExpandString($layout)
    $body | Set-Content -Path $path -Encoding UTF8
    if (!$DoNotShow.IsPresent) { Start-Process $path }

    $i++
}
