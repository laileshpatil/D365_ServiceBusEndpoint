Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser

$Connection = ''
# Chnage the org url in below connection string to your crm organization 
#$conn = Get-CrmConnection -ConnectionString "authtype=OAuth;url=https://orgb9afa5bb.crm6.dynamics.com;appid=51f81489-12ee-4a9e-aaae-a2591f45987d;redirecturi=app://58145B91-0C36-4500-8554-080854F2AC97;"
$Connection = Get-CrmConnection -ConnectionString "AuthType=Office365;Username=LaileshPatil@mywork474.onmicrosoft.com;Password=MereCrmKaP@12;Url=https://orgb9afa5bb.crm6.dynamics.com;"
Write-Host "log:Connected to crm org 11" $Connection.ConnectedOrgFriendlyName

$fetchXml =
@"
<fetch no-lock="true">
<entity name="pluginassembly">
</entity>
</fetch>
"@



Write-Host @fetchXml
$result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue

Write-host $result.CrmRecords.Count

foreach ($crmRecord in $result.CrmRecords)
{
    Write-Host "------------------------------------------------"
    Write-Host $crmRecord
}
