<#Feature to grant "Full Access" rights to a mailbox#>
<#Функция для предоставления прав "Полный доступ" для почтового ящика#>
function Full_Access_Jea{
    param (
        [Parameter(Mandatory)]
        $Mail,
        [Parameter(Mandatory)]
        $Users,
        [Parameter(Mandatory)]
        $Mail_serv,
        [Parameter(Mandatory)]
        $Path_log
    )
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName Add-MailboxPermission -DisableNameChecking

    try{
        $ErrorActionPreference = 'Stop'

        foreach ($User in $Users){
            Add-MailboxPermission -identity $Mail -User $User -AccessRights FullAcces -InheritanceType all
        }
        Remove-PSSession $Session
        $Color = "green"
        $Outcome="'Full Access' rights granted"
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total
    }
    catch {
        $Data = Get-Date
        $Exception = ($Error[0].Exception).Message
        $InvocationInfo = ($Error[0].InvocationInfo).PositionMessage
        $Value = "$Data" + " " + "$Exception"+"
        " + "$InvocationInfo"
        Add-Content -Path $Path_log -Value $Value
        Remove-PSSession $Session
        $Color = "red" 
        $Outcome = 'Выполнение не удалось. Обратитесь к системному администратору'
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total
    }
}