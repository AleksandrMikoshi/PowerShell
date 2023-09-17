<#User and mailbox creation function#>
<#Функция создания пользователи и почтового ящика#>
function Create_User {
    param (
        [Parameter(Mandatory)]
        [hashtable]$UserAttr,
        [Parameter(Mandatory)]
        [string]$DC,
        [Parameter(Mandatory)]
        [string]$Path_OU,
        [Parameter(Mandatory)]
        $Mail_serv,
        [Parameter(Mandatory)]
        $Path_log,
        [Parameter(Mandatory)]
        $DB_Mail,
        [Parameter(Mandatory)]
        $ArchiveDatabase
    )
    Set-JiraConfigServer -Server "$Jira"
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Mail_serv/PowerShell/ -Authentication Kerberos -Credential $Cred_Exch
    Import-PSSession $Session -AllowClobber -CommandName Enable-Mailbox,Get-GlobalAddressList,Update-GlobalAddressList,Get-OfflineAddressBook,Update-OfflineAddressBook,Get-AddressList,Update-AddressList -DisableNameChecking

    $LastName        = $UserAttr.LastName
    $FirstName       = $UserAttr.FirstName
    $SamAccountName  = $UserAttr.SamAccountName
    $PipeLine        = $UserAttr.PipeLine
    try{
        Add-Type -AssemblyName System.Web
        $length = Get-Random -Minimum 9 -Maximum 12
        $password = [System.Web.Security.Membership]::GeneratePassword($length, 2)
        $pass = ConvertTo-SecureString $password -AsPlainText -force
        $UserPrincipalName = $UserAttr.SamAccountName + "@m2.ru"
        $DisplayName = $UserAttr.LastName + " " + $UserAttr.FirstName
        $FullName = $UserAttr.LastName + " " + $UserAttr.FirstName + " " + $UserAttr.Initials

        $extension = @{
        'extensionAttribute1'= $UserAttr.Birth
        'extensionAttribute2' = $UserAttr.Initials
        'extensionAttribute5' = $UserAttr.Birth
        }
        $ADParameters = @{
        'Server'                = $DC
        'DisplayName'           = $DisplayName
        'GivenName'             = $UserAttr.FirstName
        'Office'                = $UserAttr.Office
        'Surname'               = $UserAttr.LastName
        'OfficePhone'           = $UserAttr.Phone
        'Department'            = $UserAttr.Department
        'Manager'               = $UserAttr.Manager
        'Title'                 = $UserAttr.JobTitle
        'UserPrincipalName'     = $UserPrincipalName
        'SamAccountName'        = $UserAttr.SamAccountName
        'Path'                  = $Path_OU
        'ChangePasswordAtLogon' = $true
        'Enabled'               = $true
        'AccountPassword'       = $pass
        'OtherAttributes'       = $extension
        }

        New-ADUser $FullName @ADParameters

        if ($UserAttr.Division -ne '') { Set-ADUser -Server $DC $SamAccountName -add @{extensionAttribute3 = "$Division" } }
        if ($UserAttr.Management -ne '') { Set-ADUser -Server $DC $SamAccountName -Replace @{extensionAttribute4 = "$Management" } }
        if ($UserAttr.Department -eq "Департамент информационных технологий") {
            $DIT = 'IT','Confluence_IT','ra_dit','logs-all-viewers'
            $DIT | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление инфраструктуры") {
            $infra_g = 'ra_admins','ra_pki','ifra_dept-1-1036737690','Confluence_Infra'
            $infra_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            } 
        if ($UserAttr.Management -eq "Управление мобильной разработки") {
            $mobile_g = 'ra_mobile','Mobile_dev-1-118242858','Developers','gitlab-dev'
            $mobile_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление обеспечения качества сервисов") {
            $qa_g = 'it_testing-1989712763','Testing','Nexus_Testing','gitlab-dev'
            $qa_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.JobTitle -eq “Разработчик 1С”) {
            $1c_g = 'local_admins_ONEC-01','BillingTeam-11178752660','1C_RemoteApp','Nexus_Users','Developers','gitlab-dev'
            $1c_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            } 
        if ($UserAttr.Management -eq "Управление серверной разработки") { 
            $back_g = 'Nexus_Backend','Nexus_Users','Developers','devsrv_dept-1-1322684850','gitlab-dev'
            $back_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление серверной разработки сервисов поиска недвижимости") { 
            $backs_g = 'Nexus_Backend','Nexus_Users','Developers','Кагал-1742951067','Backend_SPB-1-677026241','gitlab-dev'
            $backs_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление разработки интерфейсов") { 
            $front_g = 'Confluence_Frontend','Nexus_Frontend','Nexus_Users','Developers','gitlab-dev'
            $front_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление разработки интерфейсов" -And $Office -eq "Санкт-Петербург") { 
            Add-ADGroupMember Frontend_SPB-1278026215 $SamAccountName -Server $DC
            }
        if ($UserAttr.Management -eq "Управление разработки интерфейсов" -And $Office -eq "Москва (постоянный пропуск)") { 
            Add-ADGroupMember Front_msk-1713984934 $SamAccountName -Server $DC
            }
        if ($UserAttr.Department -eq "Департамент маркетинга") {
            $marketing_g = 'Marketing','Confluence_Marketing'
            $marketing_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление маркетинга") {
            Add-ADGroupMember Conf_VLot_main $SamAccountName -Server $DC
            }
        if ($UserAttr.Department -eq "Департамент по обеспечению безопасности") {
            Add-ADGroupMember 'Рассылка увольнение-1-1775093581' $SamAccountName -Server $DC
            }
        if ($UserAttr.Management -eq "Управление информационной безопасности") { 
            $itsec_g = 'soc@m2.ru-1270106079','alert-sec-1-129326235','pass_security','logs-siem-admins','local_admins_KSC','workstation_admins','Security'
            $itsec_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление экономической безопасности") { 
            $finsec_g = 'Jira_Security','owncloud_security-dep_fa','Security'
            $finsec_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Department -eq "Департамент развития сервисов для владельцев недвижимости") {
            $remont_g = 'Cognos_Users_Repair','Confluence_Legal_Statutory_Documents','ra_backoffice'
            $remont_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Department -eq "Департамент развития сервисов сделок с недвижимостью") {
            Add-ADGroupMember ra_backoffice $SamAccountName -Server $DC
            }
        if ($UserAttr.Department -eq "Департамент продаж") {
            $sales_g = 'Confluence_Sellers','ra_backoffice'
            $sales_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление по работе с банками и страховыми компаниями") { 
            $banks_g = 'Cognos_Sales','Cognos_Users_MB'
            $banks_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление по работе с застройщиками") { 
            $zastroy_g = 'Cognos_Sales','Cognos_look_SLT','Продажи_Застройщики-12079901082','Продажи_Застройщики и Риелторы-1914997088'
            $zastroy_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.JobTitle -eq “Специалист по работе с риелторами” -Or $JobTitle -eq “Специалист по телемаркетингу” -Or $JobTitle -eq “Специалист по работе с ипотечными заявками” -Or $JobTitle -eq "Специалист по поддержке клиентов" -Or $JobTitle -eq “* по обработке письменных обращений" -Or $JobTitle -eq “* контакт-центра” -Or $JobTitle -eq “Оператор КЦ”) {          
            $contact_g = 'ccteam@m2.ru-1-1838355575','Confluence_Contact','ra_backoffice'
            $contact_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление развития CRM") { 
            $crm_g = 'ra_dit','Nexus_Analytics'
            $crm_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.JobTitle -eq “Инженер данных” -Or $JobTitle -eq “Инженер данных CRM”) {
            Add-ADGroupMember Cognos_Data_Engineers $SamAccountName
            }
        if ($UserAttr.Department -eq "Департамент развития сервисов имущественных торгов") {
            $vlot_g = 'Conf_VLot_main','v-lot_users','ra_backoffice'
            $vlot_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление разработки сервисов имущественных торгов") { 
            $vlotdev_g = 'ra_dit','logs-vlot-viewer','Nexus_Vlot','Developers','IT','gitlab-dev'
            $vlotdev_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.JobTitle -eq “Менеджер по ипотечному кредитованию”) {
            $morgage_g = 'Jira_Morgage_Managers','Confluence_Sellers'
            $morgage_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление по работе с партнерами") { 
            Add-ADGroupMember op-1-721006594 $SamAccountName -Server $DC
            }
        if ($UserAttr.Management -eq "Управление по сопровождению сделок") { 
            $sdelka_g = 'ra_backoffice'
            $sdelka_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление развития мобильных сервисов") { 
            $mobprogress_g = 'ra_backoffice','ra_mobile','Property_Search','Product','Nexus_Users','Mobile_dev','Confluence_Product','Confluence_Search_Block','Confluence_Deal_Block'
            $mobprogress_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq 'Управление развития сервиса "Ипотечный брокер"') { 
            $ipbroker_g = 'ra_backoffice','banki_ru','Команда брокера-1-439681594','Поиск недвижимости-1-364314875'
            $ipbroker_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Department -eq "Департамент стратегии и финансов") {
            $finance_g = 'ra_accounter','Jira_Buy'
            $finance_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Department -eq "Департамент управления персоналом") {
            $hr_g = 'Confluence_HR','Confluence_HR_kadr_izmeneniya','Jira_HR','HR','HR_Msk'
            $hr_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление кадрового администрирования") { 
            $kadry_g = 'ra_accounter','1C_RemoteApp','Отдел кадров-11795811932','Рассылка увольнение-1-1775093581'
            $kadry_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление по подбору персонала") { 
            $podbor_g = 'Confluence_HR_poisk','Nexus_users'
            $podbor_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление обучения и развития персонала") { 
            $education_g = 'bitrix_admins','ra_backoffice'
            $education_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление внутреннего аудита") { 
            $audit_g = 'Jira_Buy','Jira_Dept.head','ra_accounter','ra_backoffice','ra_prod_owner','Cognos_Users_Call','legal-1-914345463','Confluence_Management_SLT','Confluence_Heads','Confluence_HR_kadr_izmeneniya'
            $audit_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Department -eq "Управление продуктового дизайна" -or $Department -eq "Департамент дизайна") {
            $design_g = 'ra_designers','design','Confluence_Design'
            $design_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Department -eq "Департамент учёта и отчётности") {
            $buh_g = '1C_DB_Doc','1C_Platform_install','1C_RemoteApp','Confluence_Accounting','owncloud_accounting_rw','payment-1-1225497172','ra_accounter'
            $buh_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.JobTitle -eq “Бухгалтер по работе с поставщиками”) {
            Add-ADGroupMember Jira_Buy $SamAccountName -Server $DC
            }
        if ($UserAttr.JobTitle -eq “Бухгалтер по расчету заработной платы”) {
            $buhsal_g = 'Зарплата-1-718454846','Рассылка увольнение-1-1775093581'
            $buhsal_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.JobTitle -eq “Бухгалтер по реализации” -Or $JobTitle -eq “Старший бухгалтер по реализации”) {
            $buhreal_g = 'Госпошлина','Возвраты от клиентов-11103595248','Мультибонус ДУИО-1-681805088','Онлайн сделка ДУИО-1-545271471'
            $buhreal_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Department -eq "Юридический департамент") {
            $legal_g = 'legal-1-1356913011','legal-1-914345463','Confluence_Legal'
            $legal_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq "Управление по организации документооборота") { 
            Add-ADGroupMember Jira_ED_OD $SamAccountName -Server $DC
            }
        if ($UserAttr.Management -eq "Управление по сопровождению продуктов") { 
            $legalprod_g = '1C_Platform_install','1C_DB_Doc'
            $legalprod_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq 'Управление развития сервиса "Личный кабинет клиента"') { 
            $lkprogress_g = 'Confluence_Product','Product','ra_backoffice'
            $lkprogress_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        if ($UserAttr.Management -eq 'Управление развития сервиса "Личный кабинет риелтора"') { 
            $lkprogress_g = 'Confluence_Product','Product','ra_backoffice','ЛК Ролевая Модель-1-1343470823'
            $lkprogress_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
            }
        $default_g = 'Jira_Users','Confluence_Users','owncloud_users','bitrix_users','jrnlng','ra_users'
        $default_g | ForEach-Object {Add-ADGroupMember $PSItem $SamAccountName -Server $DC}
    
        $user_creds = "Логин: $SamAccountName`nПароль: $password"
        $parameters = @{
            Fields = @{
                customfield_10641 = "$user_creds"
            }
        }
        Set-JiraIssue @parameters -Issue "$PipeLine" -Credential $Cred_Jira
    
        Enable-Mailbox -identity $UserAttr.SamAccountName -Alias $UserAttr.SamAccountName -Database $DB_Mail -DomainController $DC
        Enable-Mailbox -identity $UserAttr.SamAccountName -archive -ArchiveDatabase $ArchiveDatabase -DomainController $DC
        Get-GlobalAddressList | Update-GlobalAddressList
        Get-OfflineAddressBook | Update-OfflineAddressBook
        Get-AddressList | Update-AddressList
        Remove-PSSession $Session
        $Color = "green"
        $Outcome="
            Пользователь $LastName $FirstName создан
            Почтовый ящик '$SamAccountName@m2.ru' создан"
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