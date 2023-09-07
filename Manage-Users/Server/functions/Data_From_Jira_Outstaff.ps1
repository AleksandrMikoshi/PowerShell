<#Function of getting data from Jira and passing to further user creation#>
<#Функция получения данных из Jira и передача для дальнейшего создания пользователя#>
function Data_From_Jira_Outstaff{
    param (
        [Parameter(Mandatory)]
        [string]$PipeLine,
        [Parameter(Mandatory)]
        [string]$Jira,
        [Parameter(Mandatory)]
        [securestring]$Cred_Jira
    )   
    Set-JiraConfigServer -Server "$Jira"
    $Full_Data = Get-JiraIssue -Issue "$PipeLine" -Credential $Cred_Jira -Fields customfield_10708,customfield_10408,customfield_10506,customfield_10502,customfield_10503,customfield_11305,customfield_10501 #получение данных из Jira
    [array]$Array_City        =   $Full_Data.customfield_10501
    [string]$Manager          =   $Full_Data.customfield_10708.name
    [string]$LastName         =   ($Full_Data.customfield_10408).Split(" ")[0]
    [string]$FirstName        =   ($Full_Data.customfield_10408).Split(" ")[1]
    [string]$Initials         =   ($Full_Data.customfield_10408).Split(" ")[2]
    [string]$SamAccountName   =   ($Full_Data.customfield_10506).ToLower()
    [string]$Office           =   $Array_City.Value
    [string]$Department       =   $Full_Data.customfield_10502
    [string]$Phone            =   $Full_Data.customfield_10503
    [string]$Company          =   $Full_Data.customfield_11305

    $UserAttr = @{
        LastName = $LastName
        FirstName = $FirstName
        Initials = $Initials
        SamAccountName = $SamAccountName
        Office = $Office
        Department = $Department
        Phone = $Phone
        Manager = $Manager
        Company = $Company}
        $UserAttr
    }
     
    