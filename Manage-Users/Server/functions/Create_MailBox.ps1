<#Mailbox creation function#>
<#Функция создания почтового ящика#>
function Create_MailBox {
    param (
        [Parameter(Mandatory)]
        [hashtable]$MailAttr,
        [Parameter(Mandatory)]
        [string]$DC,
        [Parameter(Mandatory)]
        [string]$Path_Mail,
        [Parameter(Mandatory)]
        $Mail_serv,
        [Parameter(Mandatory)]
        $Path_log
    )
    $Name = $UserAttr.Name
    $UserPrincipalName = $UserAttr.UserPrincipalName
    $password_mail = $UserAttr.password_mail

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName New-Mailbox -DisableNameChecking

 
    try {
        $ErrorActionPreference = 'Stop'

        New-Mailbox -Name $Name -UserPrincipalName $UserPrincipalName -Password $password_mail - DisplayName $Name -OrganizationUnit $Path_Mail -Database DAG_DB02 -DomainController $DC
        Remove-PSSession $Session
        $Color = "green"
        $Outcome="Mailbox $Name ($UserPrincipalName) created"
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
 
