<#Manage-User - menu for command for create, edit, delete account in Active Directory end Exchange.
The data for creating users is collected from the ticket in JIRA. Use module from Atlassian company - https://atlassianps.org/docs/JiraPS/ #>
<#Manage-User — меню для команды создания, редактирования, удаления учетной записи в Active Directory и Exchange.
Данные для создания пользователей собираются из заявки в JIRA. Используйте официальный модуль от компании Atlassian - https://atlassianps.org/docs/JiraPS/ #>
function Manage-user{
    $global:DC = '*AD DC*'
    $global:Mail_serv = '*MAIL SERV*'
    $global:Jira = '*JIRA SERV*'
    $global:UserName = "login for connect to JIRA"
    $global:SecurePasswd = "Password for connect to JIRA" | ConvertTo-SecureString -AsPlainText -Force
    $global:cred = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePasswd
    Set-JiraConfigServer -Server "$Jira"

    
    function Start_menu{
$StartMenu=@"
====================================================
1`tManage account in Active Directory
2`tManage account in Exchange
    
Q`tExit
    
Select menu item or press Q to exit
"@  
        $Process = $False
        do{
            Clear-Host
            Write-Host ""
            Write-Host "Manage account" -ForegroundColor Green
            $Choice = Read-Host $StartMenu
            switch($Choice){
                "1" {
                    AD_menu
                }
                "2" {
                    EXCH_menu
                }
                "q" {
                    exit
                }
                default {
                    Write-Host ""
                    Write-Host "Error. There is no such item." -ForegroundColor Red
                    $process = $False
                }
            }
        }
        While ($Process -eq $false)
    }
    function AD_menu{
        Clear-Host
$ADMenu=@"
====================================================
1`tCreate new user
2`tCopy AD-groups
3`tReset password for user
4`tUnblock user
5`tFire an employee
    
    
B`tCome back
Q`tExit
    
Select menu item or press Q to exit
"@
        $Process = $False
        do{
            Write-Host ""
            Write-Host "Manage account in Active Directory" -ForegroundColor Green
            $Choice = Read-Host $ADMenu
            switch($Choice){ 
                "1" {
                $PipeLine = Read-Host "Enter number ticket in JIRA"
                Info_User -PipeLine $PipeLine
                Continue
                }
                "2" {
                $SourceUser = Read-Host "Specify the user from whom we copy groups"
                $TargetUser = Read-Host "Specify the user to whom we copy the groups"
                Copy_Group -SourceUser $SourceUser -TargetUser $TargetUser
                Continue
                }
                "3" {
                $User = Read-Host "Specify the user"
                Reset_Passwd -User $User
                Continue
                }
                "4" {
                $User = Read-Host "Specify the user"
                Unlock_Account -User $User
                Continue
                }
                "5" {
                $User = Read-Host "Specify the user"
                Fired_User -User $User
                Continue
                }
                "6" {#Secret room with cats :)
                    Write-Host "
                          /^--^\     /^--^\     /^--^\
                          \meow/     \meow/     \meow/
                         /      \   /      \   /      \
                        |        | |        | |        |
                         \__  __/   \__  __/   \__  __/
    |^|^|^|^|^|^|^|^|^|^|^|^\ \^|^|^|^/ /^|^|^|^|^\ \^|^|^|^|^|^|^|^|^|^|^|^|
    | | | | | | | | | | | | |\ \| | |/ /| | | | | | \ \ | | | | | | | | | | |
    ########################/ /######\ \###########/ /#######################
    | | | | | | | | | | | | \/| | | | \/| | | | | |\/ | | | | | | | | | | | |
    |_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|"
                Continue
                }
                "B"{return
                }
                "q" {
                exit
                }
                default {
                Clear-Host
                Write-Host ""
                Write-Host "Error. There is no such item." -ForegroundColor Red
                $process = $False
                }
            }
        }   
        While ($Process -eq $false)
    }
    function Info_User {
        param (
        [Parameter(Mandatory)]
        [string]$PipeLine
        )   
        process{
            Clear-Host
            $Full_Data = Get-JiraIssue -Issue "$PipeLine" -Credential $cred -Fields customfield_10708,customfield_10408,customfield_10506,customfield_10410,customfield_10502,customfield_10503,customfield_10505,customfield_10501,Key #Specify your fields!
            [array]$Full_Name = ($Full_Data.customfield_10408).Split(" ")
            [array]$Array_City = $Full_Data.customfield_10501
            [System.DateTime]$Data   =   get-date ($Full_Data.customfield_10505)
            [string]$Manager          =   $Full_Data.customfield_10708.name
            [string]$LastName         =   $Full_Name[0]
            [string]$FirstName        =   $Full_Name[1]
            [string]$SamAccountName   =   ($Full_Data.customfield_10506).ToLower()
            [string]$Office           =   $Array_City.Value
            [string]$JobTitle         =   $Full_Data.customfield_10410
            [string]$Department       =   $Full_Data.customfield_10502
            [string]$Phone            =   $Full_Data.customfield_10503
            [string]$Birth            =   Get-Date $Data -format "dd.MM.yyyy"
            [string]$Task             =   $Full_Data.Key
        $menu_item = $null
        While ($menu_item -ne "q"){
$Menu_2=@"
====================================================
1`tChange Last name         ($LastName)
2`tChange First name        ($FirstName)
3`tChange sAMAccountName    ($SamAccountName)
4`tChange city              ($Office)
5`tChange job               ($JobTitle)
6`tChange Department        ($Department)
7`tChange manager           ($Manager)
8`tChange phone number      ($Phone)
9`tChange birthsday         ($Birth)
    
B`tCome back
C`tContinue creating an account
Q`tExit
    

"@
       
        Write-Host ""
        Write-Host "Ticket $PipeLine. Check the data and if necessary - change" -ForegroundColor Green
        $Choice = Read-Host  $Menu_2
        switch ($Choice){
        "1" {
                Clear-Host
                $NewLastName = Read-Host "Enter Last name"
                $LastName = "$NewLastName"
                Continue
            }
        "2" {
                Clear-Host
                $NewFirstName = Read-Host "Enter First name"
                $FirstName = "$NewFirstName"
                Continue
            }
        "3" {
                Clear-Host
                $NewSamAccountName = Read-Host "Enter sAMAccountName"
                $SamAccountName = "$NewSamAccountName"
                Continue
            }
        "4" {
                Clear-Host
                $NewOffice = Read-Host "Enter city"
                $Office = "$NewOffice"
                Continue
            }
        "5" {
                Clear-Host
                $NewJobTitle = Read-Host "Enter job"
                $JobTitle = "$NewJobTitle"
                Continue
            }
        "6" {
                Clear-Host
                $NewDepartment = Read-Host "Enter department"
                $Department = "$NewDepartment"
                Continue
            }
        "7" {
                Clear-Host
                $NewManager = Read-Host "Enter sAMAccountName manager"
                $Manager = "$NewManager"
                Continue
            }
        "8" {
                Clear-Host
                $NewPhone = Read-Host "Enter phone number"
                $Phone = "$NewPhone"
                Continue
            }
        "9" {
                Clear-Host
                $NewBirth = Read-Host "Enter birthday"
                $Birth = "$NewBirth"
                Continue
            }
        "B" {
                return
            }
        "C" {
                Create_User -LastName $LastName -FirstName $FirstName -Initials $Initials -SamAccountName $SamAccountName -Office $Office -JobTitle $JobTitle -Department $Department -Phone $Phone -Birth $Birth -Task $Task -PipeLine $PipeLine -Manager $Manager -DC $DC
                return
            }
        "Q" {
                Exit
            }
        default {
            Write-Host ""
            Write-Host "Error. There is no such item." -ForegroundColor Red
            $menu_item = $False
            }
            }
        }
    }
}
    function Create_User {
        param (
            [Parameter(Mandatory)]
            [string]$LastName,
            [Parameter(Mandatory)]
            [string]$FirstName,
            [Parameter(Mandatory)]
            [string]$Initials,
            [Parameter(Mandatory)]
            [string]$SamAccountName,
            [Parameter(Mandatory)]
            [string]$Office,
            [Parameter(Mandatory)]
            [string]$JobTitle, 
            [Parameter(Mandatory)]
            [string]$Department,
            [Parameter(Mandatory)]
            [string]$Phone,
            [Parameter(Mandatory)]
            [string]$Birth, 
            [Parameter(Mandatory)]
            [string]$Task,
            [Parameter(Mandatory)]
            [string]$PipeLine,
            [Parameter(Mandatory)]
            [string]$Manager,
            [Parameter(Mandatory)]
            [string]$DC
        )
        process {
            $UserPrincipalName = $SamAccountName + "@example.com"
            $DisplayName = $LastName + " " + $FirstName
            $length = Get-Random -Minimum 8 -Maximum 11
            $nonAlphaChars = 1
            $password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
            New-ADUser `
            -Server $DC `
            -Name $DisplayName `
            -DisplayName $DisplayName `
            -GivenName $FirstName `
            -Office $Office `
            -Surname $LastName `
            -OfficePhone $Phone `
            -Department $Department `
            -Manager $Manager `
            -Title $JobTitle `
            -UserPrincipalName $UserPrincipalName `
            -SamAccountName $SamAccountName `
            -Path "*ENTER PATH OU=*" `
            -AccountPassword (ConvertTo-SecureString $password -AsPlainText -force) `
            -ChangePasswordAtLogon $true `
            -Enabled $true
            Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute1 = $Birth}
            Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute2 = $Initials }
            Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute15 = "$Birth" }
            Add-ADGroupMember 'Domain user' -Server $DC $SamAccountName
    
            $user_creds = "SamAccountName: $SamAccountName`nPassword: $password"
            $parameters = @{
                Fields = @{
                    customfield_10641 = "$user_creds"
                }
            }
        Set-JiraIssue @parameters -Issue "$PipeLine" -Credential $cred
    
        Enable-Mailbox -identity $SamAccountName -Alias $SamAccountName -Database '*DB Exchange*' -DomainController $DC
        Enable-Mailbox -identity $SamAccountName -archive -ArchiveDatabase '*DB Exchange_archive' -DomainController $DC
        Get-GlobalAddressList | Update-GlobalAddressList
        Get-OfflineAddressBook | Update-OfflineAddressBook
        Get-AddressList | Update-AddressList
        Write-Host "User $LastName $FirstName create" -ForegroundColor Green
        Write-Host "Email '$SamAccountName@example.com' create" -ForegroundColor Green
        return
        }
    }
    function Copy_Group {
        param (
            [Parameter(Mandatory)]
            $SourceUser,
            [Parameter(Mandatory)]
            $TargetUser
        )
        process{
            Clear-Host
            $SourceGroups = Get-ADPrincipalGroupMembership -Identity $SourceUser | Where-Object {$_.name -ne "Domain Users"}
            Add-ADPrincipalGroupMembership -Identity $TargetUser -MemberOf $sourceGroups
            Write-Host "Groups copied." -ForegroundColor Green
            return
        }
    }
    function Fired_User{
        param(
            [Parameter(Mandatory)]
            $User
            )
        process{
            Clear-Host
            $group_Member = Get-ADPrincipalGroupMembership -Identity $User | Where-Object { $_.Name -ne "Domain Users" }
            Get-ADUser -Identity $User | Move-ADObject -targetpath "*ENTER PATH OU= where to transfer account*"
            Remove-ADPrincipalGroupMembership -Identity $User -MemberOf $group_Member -Confirm:$false
            Set-Mailbox -Identity $User -HiddenFromAddressListsEnabled $true
            Set-ADUser -Identity $User -Enabled $false
            
            Disable-mailbox -Identity $User -Confirm:$false
            Get-GlobalAddressList | Update-GlobalAddressList
            Get-OfflineAddressBook | Update-OfflineAddressBook
            Get-AddressList | Update-AddressList
            Write-Host "Account $User disabled" -ForegroundColor Green
            return
        }
    }
    function Reset_Passwd{
        param(
            [Parameter(Mandatory)]
            $User
            )
        process{
            Clear-Host
            Add-Type -AssemblyName System.Web
            $length = Get-Random -Minimum 8 -Maximum 11
            $Generate = [System.Web.Security.Membership]::GeneratePassword($length, 2)
            $NewPasswd = ConvertTo-SecureString $Generate -AsPlainText -force
            Set-ADAccountPassword $User -NewPassword $NewPasswd -Reset -PassThru | Set-ADuser -ChangePasswordAtLogon $True
            Write-Host "New password:" $Generate -ForegroundColor Green
            return
        }
    }
    function Unlock_Account{
        param(
            [Parameter(Mandatory)]
            $User
            )
        process{
            $Check = Get-ADUser -Identity $User -Properties LockedOut
            Clear-Host
            if ($Check -eq $False) {
                Get-ADUser -Identity $User | Unlock-ADAccount
                Clear-host
                Write-Host "Account $User unblock" -ForegroundColor Green
            }
            else{
                Clear-host
                Write-Host "Account $User not blocked" -ForegroundColor Green
            }
            return
        }
    }
    function EXCH_menu{            
$ExchMenu=@"
====================================================
1`tCreate a mailbox
2`tGrant "Send As" rights
3`tGrant "Full Control" rights
    
    
B`tCome back
Q`tExit
        
Select menu item or press Q to exit
"@
    $Process = $False
        do{
            Write-Host ""
            Write-Host "Manage account in Exchange" -ForegroundColor Green #Заголовок меню
            $Choice = Read-Host $ExchMenu
            switch($Choice){
                "1" {
                    Clear-Host
                    $Name = Read-Host "Enter name mailbox"
                    $UserPrincipalName = Read-Host "Enter mailbox address (with domain name @example.com)"
                    $length = Get-Random -Minimum 8 -Maximum 11
                    $generate = [System.Web.Security.Membership]::GeneratePassword($length, 2)
                    $password = ConvertTo-SecureString $generate -AsPlainText -Force
                    CreateMail -name $Name -UserPrincipalName $UserPrincipalName -DC $DC -password $password
                    Continue
                }
                "2" {
                    Clear-Host
                    Write-Host "The 'Send As' permission allows you to delegate sending email from this mailbox.
                                The message will look like it was sent by the owner of the mailbox." -ForegroundColor Green
                    $Mail = Read-Host "Enter mailbox address"
                    $User = Read-Host "Enter sAMAccountName"
                    SendAs -Mail $Mail -User $User
                    Continue
                } 
                "3" {
                    Clear-Host
                    Write-Host "The Full Access permission allows the delegate to open this mailbox and act as the owner of this mailbox." -ForegroundColor Green
                    $Mail = Read-Host "Enter name mailbox"
                    $User = Read-Host "Enter sAMAccountName"
                    FullAccess -Mail $Mail -User $User
                    Continue
                }
                "B"{
                    return
                }
                "q" {
                    exit
                }
                default {
                    Clear-Host
                    Write-Host ""
                    Write-Host "Error. There is no such item." -ForegroundColor Red
                    $process = $False
                }
            }
        }
            While ($Process -eq $false)
    }
    function CreateMail{
        param (
            [Parameter(Mandatory)]
            $Name,
            [Parameter(Mandatory)]
            $UserPrincipalName,
            [Parameter(Mandatory)]
            $DC,
            [Parameter(Mandatory)]
            $password
            )
        process{
            Clear-Host
            $OU_mail = '*ENTER PATH OU=*'
            New-Mailbox -Name $Name -UserPrincipalName $UserPrincipalName -Password $password -DisplayName $Name -OrganizationalUnit $OU_mail -Database '*DB Exchange*' -DomainController $DC
            Clear-Host
            Write-Host 'Mailbox $Name ($UserPrincipalName) create' -ForegroundColor Green
            return
        } 
    }
    function SendAs{
        param (
            [Parameter(Mandatory)]
            $Mail,
            [Parameter(Mandatory)]
            $User
            )
        process{
            Add-ADPermission -Identity $Mail -User $User -AccessRights ExtendedRight -ExtendedRights "Send As"
            Clear-Host
            Write-Host "Permission 'Send As' granted" -ForegroundColor Green
            return
        }
    } 
    function FullAccess{
        param (
            [Parameter(Mandatory)]
            $Mail,
            [Parameter(Mandatory)]
            $User
            )
        process{
            Add-MailboxPermission -identity $Mail -User $User -AccessRights FullAcces -InheritanceType all
            Clear-Host
            Write-Host "Permission 'Full Acces' granted" -ForegroundColor Green
            return
        }
    } 
}