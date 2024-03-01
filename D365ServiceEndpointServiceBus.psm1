

function Get-D365Connection {
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$true, Position=1)][string]$ConnectionString
    )
 
    # Dynamics 365 requires [System.Net.SecurityProtocolType]::Tls12
    Set-Tls12SecurityProtocol
    Load-Module -Name Microsoft.Xrm.Data.PowerShell -Version 2.8.14
   
    # Reuse connection if can
    #if (!(Use-ActiveCrmConnection -Connection $conn -ConnectionString "$ConnectionString")) {
        $conn = Connect-CrmOnline -ConnectionString "$ConnectionString"
    #}
 
    if ($conn.IsReady -eq $false) {
        Write-Error "If using ClientSecret connection, please ensure the AppUser has been registered and given System Administrator role"
        throw $conn.LastCrmError
    }
 
    return $conn
}

function Get-ServiceEndpoint
(
    [parameter(Mandatory=$true)][string] $ServiceEndpointName,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
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
    <filter type="and" >
      <condition attribute="name" operator="eq" value="$ServiceEndpointName" />
    </filter>
  </entity>
</fetch>
"@
 
    Write-Host $fetchXml
   
    $response = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue

    Write-Host "fetch response: " $response
 
    $endpoint = $null
   
    if ($response.CrmRecords.Count -gt 0) {
        $endpoint = $response.CrmRecords[0]
    }
 
    return $endpoint
}

function Add-TopicServiceEndpoint
(
    [string] $ServiceEndpointName,
    [string] $ServiceBusNamespace,
    [string] $ServiceBusUrl,
    [string] $SasKeyName,
    [string] $SasKey,
    [string] $TopicName,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
   
    $existing = $null
 
    Write-Host "Checking for existing service endpoint for Topic $TopicName, ServiceEndpoint $ServiceEndpointName" -ForegroundColor Green
    if ($ServiceBusName) {
        $existing = Get-ServiceEndpoint -Connection $Connection -ServiceEndpointName $ServiceEndpointName
    }
 
    $endpoint = @{ }
    $endpoint.Add("name", $ServiceEndpointName)
    $endpoint.Add("solutionnamespace", $ServiceBusNamespace)
    $endpoint.Add("namespaceaddress", $ServiceBusUrl)
    $endpoint.Add("saskeyname", $SasKeyName)
    $endpoint.Add("saskey", $Saskey)
    $endpoint.Add("path", $TopicName)
    $endpoint.Add("authtype", (New-CrmOptionSetValue 2)) #sas key
    $endpoint.Add("contract", (New-CrmOptionSetValue 5)) #destination type = topic
    $endpoint.Add("messageformat", (New-CrmOptionSetValue 2)) #messageformat type = Json

    Write-Host endpoint
 
    if ($existing) {
        Write-Host "Existing service endpoint" -ForegroundColor Green
        #Set-CrmRecord -conn $Connection -Fields $endpoint -Id $existing.serviceendpointid -EntityLogicalName "serviceendpoint"
    } else {
        Write-Host "Creating service endpoint" -ForegroundColor Green
        #$endpoint.Add("serviceendpointid", $ServiceEndpointId)
        $id = New-CrmRecord -conn $Connection -Fields $endpoint -EntityLogicalName "serviceendpoint"
       
        Write-Host "Created with ServiceEndpointId = $id"
    }
}
 
function Remove-ServiceEndpoint
(
    [string] $ServiceEndpointName,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Checking for existing service endpoint for Topic $TopicName, ServiceEndpoint $ServiceEndpointName" -ForegroundColor Green
    if ($ServiceBusName) {
        $existing = Get-ServiceEndpoint -Connection $Connection -ServiceEndpointName $ServiceEndpointName
    }
 
    if ($existing) {
        Write-Host "Deleting service endpoint" -ForegroundColor Green
        Remove-CrmRecord -conn $Connection -CrmRecord $existing
    }
    else {
        Write-Host "Service endpoint does not exist" -ForegroundColor Green
    }
}


function Add-EntitySdkMessageProcessingStep
(
    [GUID] $ServiceEndpointId,
    [Microsoft] $MessageId,
    [string] $MessageName,
    [string] $EntityLogicalName,
    [string] $ServiceName,
    [string] $ServiceBusName,
    [string] $ServiceBusUrl,
    [string] $SasKeyName,
    [string] $SasKey,
    [string] $TopicName,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
   

    <#
    function New-SdkMessageProcessingStep ($MessageName, $EntityLogicalName, $WebhookId) {
        New-CrmRecord -conn $Connection -EntityLogicalName sdkmessageprocessingstep -Fields @{
            name = "SOME NAMING SCHEME"
            description = "$MessageName of $EntityLogicalName"
            sdkmessageid = [Microsoft.Xrm.Sdk.EntityReference]::new("sdkmessage", $MessageIds[$MessageName])
            sdkmessagefilterid = Get-SdkMessageFilterRef $MessageIds[$MessageName] $EntityLogicalName
            mode = [Microsoft.Xrm.Sdk.OptionSetValue]::new(0) # 0=Syncronous
            stage = [Microsoft.Xrm.Sdk.OptionSetValue]::new(40) # 40=Post-operation
            rank = 1
            invocationsource = [Microsoft.Xrm.Sdk.OptionSetValue]::new(0) # 0=Parent
            eventhandler = [Microsoft.Xrm.Sdk.EntityReference]::new("NOT SURE FOR WEBHOOKS", $WebhookId)
            solutionid = [Microsoft.Xrm.Sdk.EntityReference]::new("solution", $SolutionId)
        }
    }

    #>
   
 
    $sdkmessageid = Get-SdkMessage -Connection $Connection -Message $Message 

    $sdkmessagefilterid = Get-SdkMessageFilterRef -Connection $Connection -MessageId $sdkmessageid -EntityLogicalName $EntityLogicalName

   
    $existing = $null
   
    $fields= @{}
    $fields.Add("name", "Event Processing after $MessageName for Entity $EntityLogicalName")
    $fields.Add("asyncautodelete", $true) #asynautodelete 1: Yes, 0: No
    $fields.add("eventhandler", $ServiceEndpointId) #serviceEndpoint reference
    $fields.Add("invocationsource", (New-CrmOptionSetValue 0)) #Run in user Context -1:Internal, 0:Parent, 1:Child
    $fields.Add("iscustomizable", $true) #iscustomizable 1: Yes, 0: No
    $fields.Add("ishidden", $false) #ishidden 1: Yes, 0: No
    $fields.Add("mode", (New-CrmOptionSetValue 1)) #Execution mode 1:asyn , 0:syn
    $fields.Add("rank", 1) #Execution Order
    #sdkmessage
    $fields.Add("sdkmessage", $sdkmessageid) # Message Id
    $fields.Add("sdkmessagefilterid", $sdkmessagefilterid)
    $fields.Add("stage", (New-CrmOptionSetValue 40)) #State: 40 - Post operation
    $fields.Add("statecode", (New-CrmOptionSetValue 0)) #0: Enabled, 1:Disabled
    $fields.Add("supporteddeployment", (New-CrmOptionSetValue 0)) #Deployment 0: Server, 1 D365 client for Outlook, 2: Both


    if ($existing) {
        Write-Host "Updating service endpoint" -ForegroundColor Green
        Set-CrmRecord -conn $Connection -Fields $endpoint -Id $existing.serviceendpointid -EntityLogicalName "serviceendpoint"
    } else {
        Write-Host "Creating service endpoint" -ForegroundColor Green
        $endpoint.Add("serviceendpointid", $ServiceEndpointId)
        $id = New-CrmRecord -conn $Connection -Fields $endpoint -EntityLogicalName "serviceendpoint"
       
        Write-Host "Created with ServiceEndpointId = $id"
    }
 
    #SdkMessageProcessingStepId
}
 
function Get-EntitySdkMessageProcessingStep
(
    [GUID] $EntityId
)
{
   
}
 



function Get-EntitySdkMessage
(
    [string] $EntityName,
    [string] $MessageType
)
{
    return [GUID]
}
 

function Get-SdkMessage
(
    [parameter(Mandatory=$true)][string] $Message,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) 
{

    $fetchXml =
    @"
<fetch no-lock="true">
    <entity name="sdkmessage">
        <attribute name="sdkmessageid" />
    <filter type="and" >
        <condition attribute="name" operator="eq" value="$Message" />
    </filter>
    </entity>
</fetch>
"@

    Write-Host "Getting sdkMessage for " $Message
    Write-Host $fetchXml

    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue
    Write-Host $result

    return $result.CrmRecords[0].sdkmessageid
 
}

function Get-SdkMessageFilterRef 
(
    [guid]  $MessageId, 
    [string] $EntityLogicalName, 
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) 
{
 
    $fetchXml =
    @"
<fetch no-lock="true">
    <entity name="sdkmessagefilter">
        <attribute name="sdkmessagefilterid" />
        <filter>
            <condition attribute="sdkmessageid" operator="eq" value="$MessageId" />
        </filter>
        <link-entity name="entity" from="objecttypecode" to="primaryobjecttypecode">
            <filter>
                <condition attribute="logicalname" operator="eq" value="$EntityLogicalName" />
            </filter>
        </link-entity>
    </entity>
</fetch>
"@
 
    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue
 
    Write-Host $result
    return $result.CrmRecords[0].sdkmessagefilterid
}

