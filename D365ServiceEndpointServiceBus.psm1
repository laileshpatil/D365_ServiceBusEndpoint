

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
    $existing = Get-ServiceEndpoint -Connection $Connection -ServiceEndpointName $ServiceEndpointName
 
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
        return $existing.serviceendpointid
    } else {
        Write-Host "Creating service endpoint" -ForegroundColor Green
        #$endpoint.Add("serviceendpointid", $ServiceEndpointId)
        $id = New-CrmRecord -conn $Connection -Fields $endpoint -EntityLogicalName "serviceendpoint"
       
        Write-Host "Created with ServiceEndpointId = $id"

        return $id
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


function Add-EntitySdkMessageProcessingStepImage(
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection,
    [guid] $Sdkmessageprocessingstepid, 
    [string] $Message, 
    [string] $EntityLogicalName,
    [string] $Name
) 
{
    Write-Host "Processing Image for sdkmessageprocessingstepid " $Sdkmessageprocessingstepid
    $fields= @{}
    
    $fields.Add("messagepropertyname", "Target")
    $fields.Add("name", $Name)

    switch ($Message.ToLower()) {
        'create' {  
            $fields.Add("imagetype", (New-CrmOptionSetValue 1))
        }
        'update'{
            $fields.Add("imagetype", (New-CrmOptionSetValue 2))
        }
        Default {
            $fields.Add("imagetype", (New-CrmOptionSetValue 0))
        }
    }

    $fields.Add("entityalias", $EntityLogicalName)
    $fields.add("sdkmessageprocessingstepid", [Microsoft.Xrm.Sdk.EntityReference]::new("sdkmessageprocessingstep", $Sdkmessageprocessingstepid))

    $existing = Get-SdkMessageProcessingStepImage -Connection $Connection -Sdkmessageprocessingstepid $Sdkmessageprocessingstepid -Name $name

    $imageId = $null

    if ($existing) {
        Write-Host "Existing sdkmessageprocessingstepimage" -ForegroundColor Green
        $imageId = $existing.sdkmessageprocessingstepimageId
        #Set-CrmRecord -conn $Connection -Fields $endpoint -Id $fields.sdkmessageprocessingstepId -EntityLogicalName "sdkmessageprocessingstep"
    } else {
        Write-Host "Creating New sdkmessageprocessingstepimage " -ForegroundColor Green
        $imageId = New-CrmRecord -conn $Connection -Fields $fields -EntityLogicalName "sdkmessageprocessingstepimage"
       
        Write-Host "Created New sdkmessageprocessingstepimageId $imageId"
    }

    Write-host "**sdkmessageprocessingstepimageId**" $imageId
    
}


function Add-EntitySdkMessageProcessingStep
(
    [GUID] $ServiceEndpointId,
    [string] $Message,
    [string] $EntityLogicalName,
    [string] $ServiceBusName,
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
   
    
    Write-Host "getting sdkmessageid for " $Message 
    $sdkmessageid = Get-SdkMessage -Connection $Connection -Message $Message 

    if($sdkmessageid)
    {
        Write-Host "sdkmessageid found " $sdkmessageid
    }else {
        throw "sdkMessageid not found." 
    }

    Write-Host "geting sdkmessagefilterid for sdkmessageid " $sdkmessageid " Entity " $EntityLogicalName
    $sdkmessagefilterid = Get-SdkMessageFilterRef -Connection $Connection -MessageId $sdkmessageid -EntityLogicalName $EntityLogicalName

    if($sdkmessagefilterid)
    {
        Write-Host "sdkmessagefilterid found" $sdkmessagefilterid
    }else {
        throw "sdkmessagefilterid not found." 
    }

    $name = ("{0} {1}: {2} of {0}" -f $EntityLogicalName, $ServiceBusName, $Message)
    $existing = $null

    Write-Host "Checking existing Message Processing Step " $name 

    $existing = Get-EntitySdkMessageProcessingStep -Connection $Connection -sdkmessageid $sdkmessageid -sdkmessagefilterid $sdkmessagefilterid -stepName $name

    Write-Host "Processing for Step " $name
    $fields= @{}
 
    $fields.Add("name", $name)

    $fields.Add("asyncautodelete", $true) #asynautodelete 1: Yes, 0: No
    $fields.add("eventhandler", [Microsoft.Xrm.Sdk.EntityReference]::new("serviceendpoint", $ServiceEndpointId)) #serviceEndpoint reference
    $fields.Add("invocationsource", (New-CrmOptionSetValue 0)) #Run in user Context -1:Internal, 0:Parent, 1:Child
    $fields.Add("iscustomizable", $true) #iscustomizable 1: Yes, 0: No
    $fields.Add("ishidden", $false) #ishidden 1: Yes, 0: No
    $fields.Add("mode", (New-CrmOptionSetValue 1)) #Execution mode 1:asyn , 0:syn
    $fields.Add("rank", 1) #Execution Order
    #sdkmessage
    $fields.Add("sdkmessageid", [Microsoft.Xrm.Sdk.EntityReference]::new("sdkmessage", $sdkmessageid)) # Message Id
    $fields.Add("sdkmessagefilterid", [Microsoft.Xrm.Sdk.EntityReference]::new("sdkmessagefilter", $sdkmessagefilterid) )
    $fields.Add("stage", (New-CrmOptionSetValue 40)) #State: 40 - Post operation
    $fields.Add("statecode", (New-CrmOptionSetValue 0)) #0: Enabled, 1:Disabled
    $fields.Add("supporteddeployment", (New-CrmOptionSetValue 0)) #Deployment 0: Server, 1 D365 client for Outlook, 2: Both

    $id = $null

    if ($existing) {
        $id = $existing.sdkMessageProcessingStepId
        Write-Host "Existing sdkmessageprocessingstep " $id -ForegroundColor Green
        #Set-CrmRecord -conn $Connection -Fields $endpoint -Id $fields.sdkmessageprocessingstepId -EntityLogicalName "sdkmessageprocessingstep"
    } else {
        Write-Host "Creating New sdkmessageprocessingstep " -ForegroundColor Green
        $id = New-CrmRecord -conn $Connection -Fields $fields -EntityLogicalName "sdkmessageprocessingstep"
       
        Write-Host "Created with sdkmessageprocessingstepId = $id"
    }

    #Adding SdkMessageProcessingStepImage
    if ($id)
    {
        $imageName = ("Image for {0} post {1}" -f $EntityLogicalName, $Message)
        Add-EntitySdkMessageProcessingStepImage -Connection $Connection -Sdkmessageprocessingstepid $id -Message $Message -EntityLogicalName $EntityLogicalName -Name $imageName
    }

    return $id
 
    #SdkMessageProcessingStepId
}
 
function Get-EntitySdkMessageProcessingStep
(
    [guid] $sdkmessageid, 
    [guid] $sdkmessagefilterid, 
    [string] $stepName,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
)
{
      
    $fetchXml =
@"
<fetch no-lock="true">
<entity name="sdkmessageprocessingstep">
<filter>
    <condition attribute="sdkmessageid" operator="eq" value="$sdkmessageid" />
    <condition attribute="sdkmessagefilterid" operator="eq" value="$sdkmessagefilterid" />
    <condition attribute="name" operator="eq" value="$stepName" />
    
</filter>  
</entity>
</fetch>
"@

Write-Host @fetchXml
$result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue

Write-Host $result.CrmRecords[0]

Return $result.CrmRecords[0]

}
 

function Get-SdkMessage
(
    [parameter(Mandatory=$true)][string] $Message,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) 
{

    Write-Host "Getting sdkMessage for " $Message

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

    
    Write-Host $fetchXml

    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue
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

function Get-SdkMessageProcessingStepImage 
(
    [guid]  $SdkMessageProcessingStepId,  
    [string] $Name,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) 
{

    Write-Host "Retrive SdkMessageProcessingStepImage for sdkmessageprocessingstepid " $Sdkmessageprocessingstepid " Name " $Name
 
    $fetchXml =
    @"
    <fetch no-lock="true">
    <entity name="sdkmessageprocessingstepimage">
        <filter>
            <condition attribute="sdkmessageprocessingstepid" operator="eq" value="$Sdkmessageprocessingstepid" />
            <condition attribute="name" operator="eq" value="$Name" />
        </filter>
    </entity>
    </fetch>
"@

    Write-Host "Query: " $fetchXml
    
    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml -WarningAction SilentlyContinue

    Write-Host "Query Result Count: " $result.CrmRecords.Count
    #Write-Host $result.CrmRecords[0]

    $response = $null

    if ($result.CrmRecords.Count -gt 0) {
        $response = $response.CrmRecords
    }    

    return $respons
}

