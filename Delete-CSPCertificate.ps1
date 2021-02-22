#Функция удаления сертификатов с истёкшим сроком действия
[System.DateTime]$Date = (get-date).AddDays(+1) #Текущая дата +1 день, т.к. файлы сертификатов подменяем заранее
function Delete-CSPCertificate{
[Selected.System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert = Get-ChildItem Cert:\CurrentUser\My | Select-Object -Property * | Where-Object {$_.NotAfter -lt $Date} #Выбираем сертификат с истёкшим сроком
Remove-Item -Path $Cert.PSPath -deletekey #Удаляем сертификат и закрытый ключ
}
