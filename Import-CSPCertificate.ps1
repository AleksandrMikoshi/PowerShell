<#Установка сертификатов пока происходит посредством программы certmgr.exe которая поставляется совместно с КриптоПРО.#>
function Import-CSPCertificate {
    param(
        [Parameter (Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,   # Путь до расположения сертификата
        [Parameter (Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$UserLogin     # Логин пользователя
    )
    $ErrorActionPreference = "Continue"
    [string]$certmgr = 'C:\Program Files\Crypto Pro\CSP\certmgr.exe'    # Запуск certmgr.exe
    try{
        [string]$searchArray = .$certmgr -list -file "$Path\$UserLogin.cer" -verbose | Select-String -Pattern SubjKeyID    # Поиск файла сертификата. Берём из него SubjKeyID
        }
    catch{
        throw "Путь $Path\$UserLogin.cer не доступен"
        }
    [string]$searchArray = .$certmgr -list -file "$Path\$UserLogin.cer" -verbose | Select-String -Pattern SubjKeyID    # Поиск файла сертификата. Берём из него SubjKeyID
    Write-Verbose -Message $searchArray
    [string]$search = ($searchArray -split "SubjKeyID           : ",2)[1]    # Удаляем из строки SubjKeyID все строки до самого ключа
    Write-Verbose -Message $search
    [string]$Install_cert = .$certmgr -list -verbose -keyid $search | Select-String -Pattern "Serial", "Not valid before", "Not valid after"    #Поиск среди установленных сертификатов по SubjKeyID. Берём данные подписанта и срок действия(если сертификата нет, то NULL).
    #Write-Verbose -Message $Install_cert
    [string]$File_cert = .$certmgr -list -file "$Path\$UserLogin.cer" | Select-String -Pattern "Serial", "Not valid before", "Not valid after"    #Данные подписанта и срок действия в файле сертификата. 
    #Write-Verbose -Message $File_cert
    if ("$Install_cert" -notlike "$File_cert"){# Сравнение полученных данных в $Install_cert и $File_cert
        #Write-Verbose -Message "Variant INSTALL"
        $Output = .$certmgr -inst -pfx -file "$Path\$UserLogin.pfx" -pin 1234567890 -store uMy -silent    #Устанавливаем контейнер закрытого ключа
        # Write-Verbose -Message $Output
        $ContainerArray = ($Output | Select-String -Pattern "Container" | Select-Object -First 1)    #Сохраняем имя контейнера 
        #Write-Verbose -Message $ContainerArray
        [string]$Container = ($ContainerArray -split "Container           : ",2)[1]    # Удаляем из строки контейнера все строки до самого контейнера
        #Write-Verbose -Message $Container 
        .$certmgr -install -store uMy -file "$Path\$UserLogin.cer" -certificate -container "$Container" -silent -inst_to_cont}    # Установка открытой части сертификата и ссылка на закрытую часть
}
