<#Certificates are currently installed using the certmgr.exe program, which comes with CryptoPRO.#>
function Import-CSPCertificate {
    param(
        [Parameter (Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,   # Path to the location of the certificate
        [Parameter (Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$UserLogin     # User login
    )
    $ErrorActionPreference = "Continue"
    # Run certmgr.exe
    [string]$certmgr = 'C:\Program Files\Crypto Pro\CSP\certmgr.exe'    
    try{
        # Search for a certificate file. We take SubjKeyID from it
        [string]$searchArray = .$certmgr -list -file "$Path\$UserLogin.cer" -verbose | Select-String -Pattern SubjKeyID
        }
    catch{
        throw "Path $Path\$UserLogin.cer is not available"
        }
    # Search for a certificate file. We take SubjKeyID from it
    [string]$searchArray = .$certmgr -list -file "$Path\$UserLogin.cer" -verbose | Select-String -Pattern SubjKeyID   
    # Remove all lines from the SubjKeyID line up to the key itself
    [string]$search = ($searchArray -split "SubjKeyID           : ",2)[1]
    # Search among installed certificates by SubjKeyID. We take the data of the signer and the validity period (if there is no certificate, then NULL)
    [string]$Install_cert = .$certmgr -list -verbose -keyid $search | Select-String -Pattern "Serial", "Not valid before", "Not valid after"    
    # Signer data and validity period in the certificate file
    [string]$File_cert = .$certmgr -list -file "$Path\$UserLogin.cer" | Select-String -Pattern "Serial", "Not valid before", "Not valid after"     
    
    # Compare the received data in $Install_cert and $File_cert
    if ("$Install_cert" -notlike "$File_cert"){
        # Set up the private key container
        $Output = .$certmgr -inst -pfx -file "$Path\$UserLogin.pfx" -pin 1234567890 -store uMy -silent    
        # Save container name
        $ContainerArray = ($Output | Select-String -Pattern "Container" | Select-Object -First 1)
        # Remove all lines from the container line up to the container itself
        [string]$Container = ($ContainerArray -split "Container           : ",2)[1]
        # Install the public part of the certificate and link to the private part
        .$certmgr -install -store uMy -file "$Path\$UserLogin.cer" -certificate -container "$Container" -silent -inst_to_cont}    
}
