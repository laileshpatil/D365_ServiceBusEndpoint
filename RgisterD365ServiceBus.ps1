
#$ServiceEndpointName = "account sb-demo-servicebusdemoterraform-aue-dev"
#$TopicName = "account"
#$ServiceEndpointId = ""
<#
[string] $ServiceEndpointName = "account sb-demo-servicebusdemoterraform-aue-dev"
[string] $ServiceBusNamespace = "sb-demo-servicebusdemoterraform-aue-dev"
[string] $ServiceBusUrl = "sb://sb-demo-servicebusdemoterraform-aue-dev.servicebus.windows.net/"
[string] $SasKeyName = "account_sasPolicy"
[string] $SasKey = "fWQgXioRFL9APZFXWLTJuYiBAOLTHN6TM+ASbMkqq74="
[string] $TopicName ="account"

[string] $ServiceEndpointName = "contact sb-demo-servicebusdemoterraform-aue-dev"
[string] $ServiceBusNamespace = "sb-demo-servicebusdemoterraform-aue-dev"
[string] $ServiceBusUrl = "sb://sb-demo-servicebusdemoterraform-aue-dev.servicebus.windows.net/"
[string] $SasKeyName = "contact_sasPolicy"
[string] $SasKey = "f2a9fskG80ssADYt6lJoYIkQyOPudW+os+ASbH4ZTbc="
[string] $TopicName ="contact"
#>

[string] $ServiceEndpointName = "account sb-demo-servicebusdemoterraform-aue-dev"
[string] $ServiceBusNamespace = "sb-demo-servicebusdemoterraform-aue-dev"
[string] $ServiceBusUrl = "sb://sb-demo-servicebusdemoterraform-aue-dev.servicebus.windows.net/"
[string] $SasKeyName = "account_sasPolicy"
[string] $SasKey = "fWQgXioRFL9APZFXWLTJuYiBAOLTHN6TM+ASbMkqq74="
[string] $TopicName ="account"


Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser

$Connection = ''
# Chnage the org url in below connection string to your crm organization 
#$conn = Get-CrmConnection -ConnectionString "authtype=OAuth;url=https://orgb9afa5bb.crm6.dynamics.com;appid=51f81489-12ee-4a9e-aaae-a2591f45987d;redirecturi=app://58145B91-0C36-4500-8554-080854F2AC97;"
$Connection = Get-CrmConnection -ConnectionString "AuthType=Office365;Username=LaileshPatil@mywork474.onmicrosoft.com;Password=MereCrmKaP@12;Url=https://orgb9afa5bb.crm6.dynamics.com;"
Write-Host "log:Connected to crm org 11" $Connection.ConnectedOrgFriendlyName

Import-Module ".\D365ServiceEndpointServiceBus.psm1" -Force

$EntityName = "account"

$fetchXml =
@"
<fetch no-lock="true">
<entity name="sdkmessagefilter">
</entity>
</fetch>
"@

$fetchXml =
@"
<fetch no-lock="true">
<entity name="sdkmessageprocessingstep">
</entity>
</fetch>
"@

$EntityLogicalName = "Account"

$fetchXml =
@"
<fetch no-lock="true">
<entity name="sdkmessagefilter">
    <attribute name="sdkmessagefilterid" />
    <link-entity name="entity" from="objecttypecode" to="primaryobjecttypecode">
        <filter>
            <condition attribute="logicalname" operator="eq" value="$EntityLogicalName" />
        </filter>
    </link-entity>
</entity>
</fetch>
"@

$fetchXml =
@"
<fetch no-lock="true">
<entity name="sdkmessageprocessingstep">
</entity>
</fetch>
"@

$Message = "Create"
$fetchXml =
@"
<fetch no-lock="true">
<entity name="sdkmessage">
<attribute name="sdkmessageid" />
<attribute name="name" />
<filter type="and" >
<condition attribute="name" operator="eq" value="$Message" />
</filter>
</entity>
</fetch>
"@




#Write-Host @fetchXml
#$result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue

#write-host $result.CrmRecords.Count
#
<#
foreach ($crmRecord in $result.CrmRecords)
{
    Write-Host "------------------------------------------------"
    Write-Host $crmRecord
}
#>

#Write-Host  $result.CrmRecords[0]

[GUID]$sdkmessageid = [System.guid]::New("9ebdbb1b-ea3e-db11-86a7-000a3a5473e8")
$entityName = 1
$fetchXml =
@"
<fetch no-lock="true">
<entity name="sdkmessagefilter">
<attribute name="primaryobjecttypecode" />
<attribute name="sdkmessageid" />
    <filter>
        <condition attribute="sdkmessageid" operator="eq" value="$sdkmessageid" />
        <condition attribute="primaryobjecttypecode" operator="eq" value="$entityName" />
    </filter>  
    
</entity>    
</fetch>
"@

$EntityLogicalName = "account"

$fetchXml =
@"
<fetch no-lock="true">
<entity name="sdkmessagefilter">
    <filter>
        <condition attribute="sdkmessageid" operator="eq" value="$sdkmessageid" />
    </filter>
    <link-entity name="entity" from="objecttypecode" to="primaryobjecttypecode">
        <filter>
            <condition attribute="logicalname" operator="eq" value="$EntityLogicalName" />
        </filter>
    </link-entity>
</entity>
</fetch>
"@





try {

    $sdkmessageid = Get-SdkMessage -Connection $Connection -Message $Message 

    Write-Host $sdkmessageid

    $sdkmessagefilterid = Get-SdkMessageFilterRef -Connection $Connection -MessageId $sdkmessageid -EntityLogicalName $EntityLogicalName

    Write-Host $sdkmessagefilterid
   
    $fetchXml =
@"
<fetch no-lock="true">
<entity name="sdkmessageprocessingstep">
<link-entity name="plugingtype" from="plugintypeid" to="plugintypeid">
</link-entity>
<filter>
    <condition attribute="sdkmessageid" operator="eq" value="$sdkmessageid" />
    <condition attribute="sdkmessagefilterid" operator="eq" value="$sdkmessagefilterid" />
</filter>  
</entity>
</fetch>
"@

Write-Host @fetchXml
$result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue

Write-host $result.CrmRecords.Count


foreach ($crmRecord in $result.CrmRecords)
{
    Write-Host "------------------------------------------------"
    Write-Host $crmRecord.name
    Write-Host $crmRecord
}

    #Add-TopicServiceEndpoint -Connection $Connection -ServiceEndpointName $ServiceEndpointName -ServiceBusNamespace $ServiceBusNamespace -ServiceBusUrl $ServiceBusUrl -TopicName $TopicName -SasKeyName $SasKeyName -SasKey $SasKey
 
    #Remove-ServiceEndpoint -Connection $Connection -ServiceEndpointName $ServiceEndpointName
    exit 0
} catch {
    Write-Error $_.Exception.Message
    exit 1  
}






<#

Write-Host @fetchXml
$result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue

write-host $result.CrmRecords.Count
Write-Host  $result.CrmRecords[0].sdkmessagefilterid
Write-Host  $result.CrmRecords[0].primaryobjecttypecode
Write-Host  $result.CrmRecords[0].primaryobjecttypecode_Property.Value

#Write-Host "eventhandler_Property: "  $result.CrmRecords[0].sdkmessageprocessingstepid_Property
#Write-Host "eventhandler_Property: "  $result.CrmRecords[0].sdkmessageprocessingstepid
#Write-Host "eventhandler_Property: "  $result.CrmRecords[0].sdkmessagefilterid_Property
#Write-Host "eventhandler_Property: " $result.CrmRecords[0].organizationid_Property


  if (-not $ConnectionString) {
        $ConnectionString = $env:ConnectionString
    }
   
    if (-not $ServiceBusSasKey) {
        $ServiceBusSasKey = $env:D365_SERVICE_BUS_KEY
    }

$existing = $null
 
Write-Host "Checking for existing service endpoint for Topic $TopicName, ServiceEndpointName $ServiceEndpointName" -ForegroundColor Green
if ($ServiceEndpointName) {
    $existing = Get-ServiceEndpoint -Connection $Connection -ServiceEndpointName $ServiceEndpointName
}

Write-Host $existing

$fetchXml =
@"
<fetch>
<entity name="serviceendpoint">
<attribute name="serviceendpointid" />
<attribute name="name" />
<attribute name="path" />
<attribute name="solutionnamespace" />
<attribute name="namespaceaddress" />
<attribute name="saskeyname" />
<attribute name="authtype" />
<attribute name="contract" />
</entity>
</fetch>
"@

Write-Host $fetchXml

#serviceendpointid=06eae3d7-1bd5-ee11-904c-000d3ad18515
#name=account sb-demo-servicebusdemoterraform-aue-dev

$response = Get-CrmRecordsByFetch -conn $conn -Fetch $fetchXml -WarningAction   SilentlyContinue

Write-Host $response.CrmRecords[0]
Write-Host $response.CrmRecords[1]

$endpoint = $null

Write-Host $response.CrmRecords.Count

#$endpoint = $response.CrmRecords[0]


Write-Host $endpoint

#>