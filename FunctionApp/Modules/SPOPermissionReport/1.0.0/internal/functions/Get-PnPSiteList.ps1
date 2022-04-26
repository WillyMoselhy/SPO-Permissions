#Connect to Admin Center
$Cred = Get-Credential
Connect-PnPOnline -Url $TenantAdminURL -Credentials $Cred

#Get All Site collections - Exclude: Seach Center, Mysite Host, App Catalog, Content Type Hub, eDiscovery and Bot Sites
$SitesCollections = Get-PnPTenantSite | Where -Property Template -NotIn ("SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")

#Loop through each site collection
ForEach($Site in $SitesCollections)
{
    #Connect to site collection
    $SiteConn = Connect-PnPOnline -Url $Site.Url -Credentials $Cred
    Write-host "Generating Report for Site:"$Site.Url
    Write-Host  "Generating Report for Site:$($Site.Url)"

    #Call the Function for site collection
    $ReportFile = "C:\Temp\$($Site.URL.Replace('https://','').Replace('/','_')).CSV"
    Generate-PnPSitePermissionRpt -SiteURL $Site.URL -ReportFile $ReportFile -Recursive

    Disconnect-PnPOnline -Connection $SiteConn
}