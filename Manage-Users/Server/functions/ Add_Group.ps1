function Add_Group {
    param (
        [Parameter(Mandatory)]
        [hashtable]$UserAttr,
        [Parameter(Mandatory)]
        [string]$DC,
        [Parameter(Mandatory)]
        $Path_log
    )
    $SamAccountName  = $UserAttr.SamAccountName

    try{
        $default_g = $Group
        $default_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
    
        $Color = "green"
        $Outcome = "✓ Adding groups to Active Directory"
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
        $Outcome = '× Adding groups to Active Directory. Contact your system administrator'
        $Total = @{
            Color = $Color
            Outcome = $Outcome
        }
        $Total
    }
} 
