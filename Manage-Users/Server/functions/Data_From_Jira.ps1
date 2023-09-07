<#Function of getting data from Jira and passing to further user creation#>
<#Функция получения данных из Jira и передача для дальнейшего создания пользователя#>
function Data_From_Jira{
    param (
        [Parameter(Mandatory)]
        [string]$PipeLine
    )   
    Set-JiraConfigServer -Server "$Jira"
    $Full_Data = Get-JiraIssue -Issue "$PipeLine" -Credential $Cred_Jira -Fields customfield_10409,customfield_10708,customfield_10408,customfield_10506,customfield_10410,customfield_10502,customfield_10503,customfield_10505,customfield_10501,customfield_10700,Key
    [array]$Array_City               =   ($Full_Data.customfield_10501)
    [System.DateTime]$Data           =   get-date ($Full_Data.customfield_10505)
    [string]$Global:Manager          =   $Full_Data.customfield_10708.name
    [string]$Global:LastName         =   ($Full_Data.customfield_10408).Split(" ")[0]
    [string]$Global:FirstName        =   ($Full_Data.customfield_10408).Split(" ")[1]
    [string]$Global:Initials         =   ($Full_Data.customfield_10408).Split(" ")[2]
    [string]$Global:SamAccountName   =   ($Full_Data.customfield_10506).ToLower()
    [string]$Global:Office           =   $Array_City.Value
    [string]$Global:JobTitle         =   $Full_Data.customfield_10410
    [string]$Global:Department       =   $Full_Data.customfield_10502
    [string]$Global:Management       =   $Full_Data.customfield_10700
    [string]$Global:Phone            =   $Full_Data.customfield_10503
    [string]$Global:Birth            =   Get-Date $Data -format "dd.MM.yyyy"
    [string]$Global:Task             =   $Full_Data.Key
    [string]$Global:Division         =   $Full_Data.customfield_10409

 
    $UserAttr = @{
        LastName = $LastName
        FirstName = $FirstName
        Initials = $Initials
        SamAccountName = $SamAccountName
        Office = $Office
        JobTitle = $JobTitle
        Department = $Department
        Management = $Management
        Division = $Division
        Phone = $Phone
        Birth = $Birth
        Task = $Task
        PipeLine = $PipeLine
        Manager = $Manager}
        $UserAttr
}    