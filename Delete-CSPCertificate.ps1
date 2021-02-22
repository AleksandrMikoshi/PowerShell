function Delete-CSPCertificate{
$Cert = Get-ChildItem Cert:\CurrentUser\My | select -Property * | Where-Object {$_.NotAfter -lt $date} #Выбираем сертификат с истёкшим сроком
Remove-Item -Path $Cert.PSPath -deletekey #Удаляем сертификат и закрытый ключ
}
