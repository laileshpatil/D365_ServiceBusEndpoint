
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
[string] $EntityLogicalName = "account"


Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser

$Connection = ''
# Chnage the org url in below connection string to your crm organization 
#$conn = Get-CrmConnection -ConnectionString "authtype=OAuth;url=https://orgb9afa5bb.crm6.dynamics.com;appid=51f81489-12ee-4a9e-aaae-a2591f45987d;redirecturi=app://58145B91-0C36-4500-8554-080854F2AC97;"
$Connection = Get-CrmConnection -ConnectionString "AuthType=Office365;Username=LaileshPatil@mywork474.onmicrosoft.com;Password=MereCrmKaP@12;Url=https://orgb9afa5bb.crm6.dynamics.com;"
Write-Host "log:Connected to crm org 11" $Connection.ConnectedOrgFriendlyName

Import-Module ".\D365ServiceEndpointServiceBus.psm1" -Force


try {

    #Adding Service Endpint If not exists
    $ServiceEndpointId = $null
    $ServiceEndpointId = Add-TopicServiceEndpoint -Connection $Connection -ServiceEndpointName $ServiceEndpointName -ServiceBusNamespace $ServiceBusNamespace -ServiceBusUrl $ServiceBusUrl -TopicName $TopicName -SasKeyName $SasKeyName -SasKey $SasKey
 
    if ($ServiceEndpointId){

        Write-Host "Creating Message Processing Steps for ServiceEndpointId " $ServiceEndpointId

        $message ="Update"
        $sdkMessageProcessingStepId =  Add-EntitySdkMessageProcessingStep -Connection $Connection -ServiceEndpointId $ServiceEndpointId -EntityLogicalName $EntityLogicalName -Message $message -ServiceBusName $ServiceBusNamespace  

        if ($sdkMessageProcessingStepId)
        {
            Write-Host "sdkMessageProcessingStepId Created : " $sdkMessageProcessingStepId
        }else{
            Write-Host "sdkMessageProcessingStepId NOT created"
        }
        
    }else
    {
        Write-Host "No Service Endpoint Reference found or created"
    }

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
