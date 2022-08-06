#Function to remove expired certificates

#Current date +1 day we replace the certificate files in advance
[System.DateTime]$Date = (get-date).AddDays(+1)

function Delete-CSPCertificate{
# Select an expired certificate
[Selected.System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert = Get-ChildItem Cert:\CurrentUser\My | Select-Object -Property * | Where-Object {$_.NotAfter -lt $Date}
# Delete certificate and private key
Remove-Item -Path $Cert.PSPath -deletekey 
}
