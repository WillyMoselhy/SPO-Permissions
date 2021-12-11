function Start-SPOPermissionCollection {
    [CmdletBinding()]
    param (

    )

    #Import Required Modules
    Write-Host "Hello World!"
    #Connect to PNP and GraphAPI

    $LoginInfo = [PSCustomObject]@{
        TenantID        = '1aeaebf6-dfc4-49c8-a843-cc2b8d54a9b1'
        TenantName      = 'm365x252065'
        AppID           = '9ce25227-4018-427e-8f8d-cbc3c0d19657'
        CertificatePath = 'C:\temp\PnP Rocks2.pfx' #This can be EncodedBase64
    }


    $Cert = new-object security.cryptography.x509certificates.x509certificate2 -ArgumentList $LoginInfo.CertificatePath
    write-host "Cert Converted"
    ${env:msalps.dll.lenientLoading} = $true # Continue Module Import # This is to avoid assembly warning
    $script:MSALToken = Get-MsalToken -ClientId 9ce25227-4018-427e-8f8d-cbc3c0d19657 -ClientCertificate $cert -TenantId 1aeaebf6-dfc4-49c8-a843-cc2b8d54a9b1 -ForceRefresh
    Connect-PnPOnline -Url "https://$($LoginInfo.TenantName).Sharepoint.com" -ClientId $LoginInfo.AppID -Tenant "$($LoginInfo.TenantName).OnMicrosoft.com" -CertificatePath $LoginInfo.CertificatePath -ErrorAction Stop
    write-host "Connected to PNP"

    Write-Host $script:MSALToken.AccessToken
    Write-Host "I've got the POWER!"



    #Collect Report Data
}