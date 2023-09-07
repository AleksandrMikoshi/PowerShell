<#The function of disabling "full access" rights for the mailbox#>
<#Функция отключения прав "полный доступ" для почтового ящика#>
function Remove_FullAccess{
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
    Import-PSSession $Session -AllowClobber -CommandName Remove-MailboxPermission -DisableNameChecking

    try{
        $ErrorActionPreference = 'Stop'

        foreach ($User in $Users){
            Remove-MailboxPermission -identity $Mail -User $User -AccessRights FullAcces -InheritanceType all -Confirm:$false
        }
        Remove-PSSession $Session
        $Color = "green"
        $Outcome="'Full Access' rights disabled"
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
        $Outcome = 'Execution failed. Contact your system administrator'
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total 
    }
}