#Запуск функции для установки личных сертификатов
$Path = "Z:\" #Путь расположения сертификатов
$UserLogin = $env:UserName #Логин пользователя (для установки ЛИЧНЫХ сертификатов)
$Function = "Z:\Scripts\Установка сертификатов\Import-CSPCertificate.ps1" #Путь расположения функции
. $Function #Добавление функции
Import-CSPCertificate -Path $Path -UserLogin $UserLogin #Запуск функции с необходимыми данными
