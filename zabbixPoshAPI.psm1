function getJSON($url = $ApiURL, $object) {
    try {
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($object)
        $web = [System.Net.WebRequest]::Create($url)
        $web.Method = “POST”
        $web.ContentLength = $bytes.Length
        $web.ContentType = “application/json”
        $stream = $web.GetRequestStream()
        $stream.Write($bytes,0,$bytes.Length)
        $stream.close()
        $reader = New-Object System.IO.Streamreader -ArgumentList $web.GetResponse().GetResponseStream()
        return $reader.ReadToEnd()| ConvertFrom-Json
        $reader.Close()
    } catch {
        $err_msg = "ERROR: [getJSON] Unable to process request. Make sure API server is available. Exception details:`nException: $($_.Exception)`nLine: $($_.InvocationInfo.ScriptLineNumber)`nOffset: $($_.InvocationInfo.OffsetInLine)"
        #sendMail $errors_emails "[$server] Error from $scriptname" $err_msg
        write-host ($err_msg) -ForeGroundColor Red; Break
    }
} 
 
function evaluate-JSON {
	<#
	.Synopsis
		Send a JSON object to a Web server and evaluates the response
	.Example
		evaluate-JSON -jsonApiUrl "http://mywebservice/api" -jsonObj $objHost -errorMsg "[get-ZabbixHost] : Unable to retrieve hosts, aborting." -noDataMsg "[get-ZabbixHost] : No host matching the description"
	.Parameter jsonApiUrl
		URL of the JSON web service
	.Parameter jsonObj
		JSON request object that will be sent to Web Server
	.Parameter errorMsg
		Message to be displayed in -verbose mode if an error occured, should allow to easily determine the function and the problem
	.Parameter noDataMsg
		Message to be displayed in -verbose mode if the request is correct but did not return data
	.Notes
		NAME: evaluate-JSON
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$True,
		HelpMessage="URL of the JSON web service")]
		[string]$jsonApiUrl
		,
		[Parameter(Mandatory=$True,
		HelpMessage="JSON object to evaluate")]
		[object]$jsonObj
		,
		[Parameter(Mandatory=$False,
		HelpMessage="Message to display is error is encountered")]
		[String]$errorMsg
		,
		[Parameter(Mandatory=$False,
		HelpMessage="Message to display is no data is returned")]
		[String]$noDataMsg
	)
	Process {
		$getJSON = getJSON $jsonApiURL $jsonObj
		if ($getJSON.error){
			$error =$getJSON.error.data
			Write-Verbose "$errorMsg. Error message is $error "
			return $False
		}
		elseif ($getJSON.result.count -eq 0) {
				Write-Verbose "$noDataMsg "
				return $false
			}
		else {
				return $getJSON
			}
	}
}
function connect-Zabbix {
	[cmdletbinding()]
	Param (
		[Parameter(Mandatory=$True
		,HelpMessage="Zabbix credentials, must have API and dashboard access")]
		[PSCredential]$zabbixCredential
		,
		[Parameter(Mandatory=$True
		,HelpMessage="Zabbix API URL, usually http://zabbix-server-name/zabbix/api_jsonrpc.php")]
		[string]$zabbixApiURL
	)
	Process {
		$zabbixApiPasswd = $zabbixCredential.getNetworkCredential().password
		$zabbixApiUser = $zabbixCredential.userName
		$global:session = get-ZabbixSession $zabbixApiUser $zabbixApiPasswd $zabbixApiURL
		if (!$session.result) {
			Write-verbose "[connect-Zabbix] : Unable to connect to Zabbix, aborting"
			Remove-Variable session -Scope global
			return $false
		}	
		else {
			Write-verbose "[connect-Zabbix] : Connection to Zabbix is successfull"
			$zabbixAPIVersion = get-zabbixAPIInfo -zabbixApiURL $zabbixApiURL
			$session | Add-Member -MemberType NoteProperty -Name "zabbixAPIVersion" -Value $zabbixAPIVersion.result
			$session | Add-Member -MemberType NoteProperty -Name "zabbixApiURL" -Value $zabbixApiURL
			return $session
		}
	}
}

function get-zabbixAPIInfo {
	[cmdletbinding()]
	Param (
		[Parameter(Mandatory=$True
		,HelpMessage="Zabbix API URL, usually http://zabbix-server-name/zabbix/api_jsonrpc.php")]
		[string]$zabbixApiURL
	)
	Process {
	#construct the JSON object
		$objAPIInfo = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'apiinfo.version' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	

		#return $objGraphItem
		
		#make the request and evaluate it
		evaluate-JSON -jsonApiUrl $zabbixApiURL -jsonObj $objAPIInfo -errorMsg "[get-zabbixAPIInfo] : Unable to retrieve API info, aborting." -noDataMsg "[get-zabbixAPIInfo] : No result returned"
	}
}

#Function to setup session and return.
function get-ZabbixSession($zabbixApiUser, $zabbixApiPasswd, $zabbixApiURL) {
    #Create authentication JSON object using ConvertTo-JSON
    $objAuth = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
    Add-Member -PassThru NoteProperty method ‘user.login’ |
    Add-Member -PassThru NoteProperty params @{user=$zabbixApiUser;password=$zabbixApiPasswd} |
    Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 
    return getJSON $zabbixApiURL $objAuth
}

function get-ZabbixUser { 
	<#
	.Synopsis
		This function retrieves all users or a single user, depending if the -username parameter is used
	.Example
		get-ZabbixUser 
		get-ZabbixUser -userName smorand
	.Parameter userName
		provide a users alias to limit the scope of the search, not mandatory
	.Notes
		NAME: get-ZabbixUser
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$False
		,HelpMessage="Provide a user alias to limit the scope of the search")]
		[String]$userName
	)
	Process {
		if ($userName) {
			$filter= @{"alias" = "$userName"}			
			$objUser = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
			Add-Member -PassThru NoteProperty method 'user.get' |
			Add-Member -PassThru NoteProperty params @{output="extend";filter=$filter} |
			Add-Member -PassThru NoteProperty auth $session.result |
			Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		}
		else {
			$objUser = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
			Add-Member -PassThru NoteProperty method 'user.get' |
			Add-Member -PassThru NoteProperty params @{output="extend"} |
			Add-Member -PassThru NoteProperty auth $session.result |
			Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		}
		
		return getJSON $session.zabbixApiURL $objUser
		
	}
}

function get-ZabbixHost { 
	<#
	.Synopsis
		Retrieves Zabbix hosts based on several search options
	.Example
		get-ZabbixHost
		get-ZabbixHost -hostName "myhost"
		get-ZabbixHost -hostGroupName "mygroup"
		get-ZabbixHost -hostPattern "bdd"
		get-ZabbixHost -hostGroupName "linuxgroup" -hostPattern "zabbixserver"
	.Parameter hostName
		provide a host name to limit the scope of the search, not mandatory
	.Parameter hostGroupName
		provide a host group name pattern to limit the scope of the search to that group, not compatible with -hostname
	.Parameter hostPattern
		Limit the scope of the search to hosts with names that match the pattern, not compatible with -hostname, not mandatory
	.Parameter short
		Returns only host ids
	.Parameter sort
		sort by host name
	.Parameter selectGroup
		Return the host groups that the host belongs to in the groups property
	.Notes
		NAME: get-ZabbixHost
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(ParameterSetName="hostname",
		Mandatory=$False,
		HelpMessage="Provide a host name to limit the scope of the search")]
		[String]$hostName
		,
		[Parameter(ParameterSetName="hostgroup",
		Mandatory=$False,
		HelpMessage="provide a host name pattern to limit the scope of the search to that pattern.Could be a list sperated by ,  not compatible with -hostName")]
		[String] $hostPattern
		,
		[Parameter(ParameterSetName="hostgroup",
		Mandatory=$False,
		HelpMessage="provide a host group name pattern to limit the scope of the search to that group, only compatible with host pattern")]
		[String] $hostGroupName
		,
		[Parameter(Mandatory=$False,
		HelpMessage="Return the host groups that the host belongs to in the groups property, not mandatory")]
		[Switch] $selectGroups
		,
		[Parameter(Mandatory=$False,
		HelpMessage="short version, returns only hostids. Usefull because JSON input is limited to 2MB in size... :( , not mandatory")]
		[Switch] $short
		,
		[Parameter(Mandatory=$False,
		HelpMessage="sort by host name, not mandatory")]
		[Switch] $sort
		
	)
	Process {
		#construct the params
		$params=@{}
		
		#construct the "groupids param"
		if ($hostGroupName) {
			#get the group name id and check that there is only one! (i dont know how to use several ids yet...)
			$groupid=get-ZabbixHostGroup -hostGroupName $hostGroupName
			if  ($groupid) {
				#if only one group is returned we're good
				if($groupid.result.count -eq 1) {
					$params.add("groupids",$groupid.result.groupid)
				}
				#if more than 1, we can't use that
				else {
					Write-verbose "[get-ZabbixHost] : More than one group returned, aborting. Number of groups returned : $groupsids.result.count"
					return $false
				}
			}
			else {
				return $false
			}
		}
		#construct the "search description" param
		if ($hostPattern) {
			$search=@{}
			$search.add("host", $hostPattern)
			$params.add("search",$search)
		}
		#construct the "filter host" param
		if ($hostName) {
			$filter=@{}
			$filter= @{"host" = "$hostName"}
			$params.add("filter",$filter)
		}
		#finish the params
		if ($short) {$params.add("output", "shorten")}
		else {$params.add("output", "extend") }
		
		if ($sort) {$params.add("sortfield","host") }
		
		if ($selectGroups) {$params.add("select_groups","refer")}
		
		#construct the JSON object
		$objHost = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'host.get' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		
		#return $objHost
		
		#make the request and evaluate it
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objHost -errorMsg "[get-ZabbixHost] : Unable to retrieve hosts, aborting." -noDataMsg "[get-ZabbixHost] : No host matching the description"
	}
}

function get-ZabbixHostGroup { 
	<#
	.Synopsis
		This function retrieves all host groups or a single host group, depending if the -hostGroupName parameter is used
	.Example
		get-ZabbixHostGroup 
		get-ZabbixHostGroup -hostGroupName smorand
	.Parameter hostGroupName
		provide a host group name to limit the scope of the search, not mandatory
	.Notes
		NAME: get-ZabbixHostGroup
		AUTHOR: Simon Morand (MBVSI)
	#>
	Param(
		[Parameter(Mandatory=$False
		,HelpMessage="Provide a host group name to limit the scope of the search")]
		[String]$hostGroupName
	)
	Process {
		if ($hostGroupName) {
			$filter= @{"name" = "$hostGroupName"}			
			$objHostGroup = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
			Add-Member -PassThru NoteProperty method 'hostgroup.get' |
			Add-Member -PassThru NoteProperty params @{output="extend";filter=$filter} |
			Add-Member -PassThru NoteProperty auth $session.result |
			Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		}
		else {
			$objHostGroup = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
			Add-Member -PassThru NoteProperty method 'hostgroup.get' |
			Add-Member -PassThru NoteProperty params @{output="extend"} |
			Add-Member -PassThru NoteProperty auth $session.result |
			Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		}
		
		$getJSON = getJSON $session.zabbixApiURL $objHostGroup
		if (!$getJSON.result){
			Write-Verbose "[get-ZabbixHostGroup] : Unable to retrieve host group $hostgroupname, aborting"
			return $false
		}
		else {return $getJSON }
		
	}
}

#### Here are a few made by JP #####
function Export-ZabbixHost { 
	<#
	.Synopsis
		This function export host XML configuration
	.Example
		export-ZabbixHost 
		export-ZabbixHost -hostID 10161
	.Parameter hostID
		ID number of host, can be retreived by Get-ZabbixHost -HostName command
	.Notes
		NAME: Export-ZabbixHost
		AUTHOR: JanPaul Klompmaker
	#>
	Param(
		[Parameter(Mandatory=$True
		,HelpMessage="Provide a host id")]
		[String]$hostID
		,
		[Parameter(Mandatory=$False
		,HelpMessage="Output file")]
		[String]$output
	)
	Process {
		if ($hostID) {
			$filter= @{"hosts" = $hostID}		
			$objHostID = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
			Add-Member -PassThru NoteProperty method 'configuration.export' |
			Add-Member -PassThru NoteProperty params @{format="xml";options=$filter} |
			Add-Member -PassThru NoteProperty auth $session.result |
			Add-Member -PassThru NoteProperty id '1') | ConvertTo-Json 	
		}	
		$export = evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objHostID -errorMsg "[get-ZabbixHostbyID] : Unable to retrieve host group $objHostID, aborting." -noDataMsg "[get-ZabbixHostbyID] : No host group matching the description"
		if ($output) {
			$output = [xml]$export.result 
			$output.zabbix_export | Out-File $output  # doesnt work yet?!?!?!?!
		} else {
			[xml]$export.result
		}
	}
}

function get-ZabbixTemplate { 
	<#
	.Synopsis
		This function retrieves all templates or a singletemplate, depending if the -templateName parameter is used
	.Example
		get-ZabbixTemplate 
		get-ZabbixTemplate -templateName Linux-Server
	.Parameter templateName
		provide a host group name to limit the scope of the search, not mandatory
	.Notes
		NAME: get-ZabbixTemplate
		AUTHOR: Simon Morand (MBVSI)
	#>
	Param(
		[Parameter(Mandatory=$False
		,HelpMessage="Provide a host group name to limit the scope of the search")]
		[String]$templateName
	)
	Process {
		if ($templateName) {
			$filter= @{"host" = "$templateName"}			
			$objTemplate = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
			Add-Member -PassThru NoteProperty method 'template.get' |
			Add-Member -PassThru NoteProperty params @{output="extend";filter=$filter} |
			Add-Member -PassThru NoteProperty auth $session.result |
			Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		}
		else {
			$objTemplate = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
			Add-Member -PassThru NoteProperty method 'template.get' |
			Add-Member -PassThru NoteProperty params @{output="extend"} |
			Add-Member -PassThru NoteProperty auth $session.result |
			Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		}
		
		$getJSON = getJSON $session.zabbixApiURL $objTemplate
		if (!$getJSON.result){
			Write-Verbose "[get-ZabbixTemplate] : Unable to retrieve host group $templateName, aborting"
			return $false
		}
		else {return $getJSON }
		
	}
}

function create-ZabbixHost {
	<#
	.Synopsis
		This function creates a host in Zabbix
	.Example
		create-ZabbixHost -hostName myhost -hostFQDN myhost.lan -hostGroupName Linux-Group -templateName Template-Linux -proxyName myproxy
	.Parameter hostName
		Name of the host to be created, mandatory
	.Parameter hostFQDN
		FQDN of the host to be created, mandatory
	.Parameter hostGroupName
		provide a host group name to limit the scope of the search, mandatory
	.Parameter templateName
		provide a host group name to limit the scope of the search, mandatory
	.Parameter proxyName
		Proxy that will monitor the host, mandatory
	.Notes
		NAME: create-ZabbixHost
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$true,
		HelpMessage="Name of the host to be created")]
		[String] $hostName
		,
		[Parameter(Mandatory=$true,
		HelpMessage="FQDN of the host to be created")]
		[String] $hostFQDN
		,
		[Parameter(Mandatory=$true,
		HelpMessage="Host Group in which to place the host")]
		[String] $hostGroupName
		,
		[Parameter(Mandatory=$true,
		HelpMessage="Template to link the host with")]
		[String] $templateName
		,
		[Parameter(Mandatory=$true,
		HelpMessage="Proxy that will monitor the host")]
		[String] $proxyName
	)	
	Process
	{
		#retrieve proxy info
		$proxy = get-ZabbixProxy -proxyName $proxyName
		if (!$proxy) {
			Write-InRed "Unable to retrieve proxy ID, aborting"
			return $false
		}
		$proxyID = $proxy.result.proxyid
		$proxy = @{"proxyid" = "$proxyID"}
		
		#retrieve template info
		$template = get-ZabbixTemplate -templateName $templateName
		if (!$template) {
			Write-InRed "Unable to retrieve template ID, aborting"
			return $false
		}
		$templateID = $template.result.hostid
		$templates = @{"templateid" = "$templateID"}
		
		#retrieve host group info
		$hostgroup = get-ZabbixHostGroup -hostGroupName $hostGroupName
		if (!$hostgroup) {
			Write-InRed "Unable to retrieve host group ID, aborting"
			return $false
		}
		$hostGroupID = $hostgroup.result.groupid
		$groups = @{"groupid" = "$hostGroupID"}
		
		#create the json object
		$objHost = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'host.create' |
		Add-Member -PassThru NoteProperty params @{host=“$hostName”;dns="$hostFQDN";groups=$groups;templates=$templates;proxy_hostid=$proxyID} |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 
		
		#make the request and evaluate it
		$getJSON = getJSON $session.zabbixApiURL $objHost
		if (!$getJSON.result){
			$Error =$getJSON.error.data
			Write-InRed "Unable to create host $hostName, aborting. Error message is : $error"
			return $false
		}
		else {
			Write-Verbose "[create-ZabbixHost] : Host $hostName created successfully!"
			return $true
		}
				
	}
}

function get-ZabbixProxy { 
	<#
	.Synopsis
		This function retrieves all proxy or a single one, depending if the -proxyName parameter is used
	.Example
		get-ZabbixProxy 
		get-ZabbixProxy -proxyName zabprox10
	.Parameter proxyName
		provide a proxy name to limit the scope of the search, not mandatory
	.Notes
		NAME: get-ZabbixProxy
		AUTHOR: Simon Morand (MBVSI)
	#>
	Param(
		[Parameter(Mandatory=$False
		,HelpMessage="Provide a host proxy name to limit the scope of the search")]
		[String]$proxyName
	)
	Process {
		if ($proxyName) {
			$search= @{"host" = "$proxyName"}			
			$objProxy = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
			Add-Member -PassThru NoteProperty method 'proxy.get' |
			Add-Member -PassThru NoteProperty params @{output="extend";search=$search} |
			Add-Member -PassThru NoteProperty auth $session.result |
			Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		}
		else {
			$objProxy = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
			Add-Member -PassThru NoteProperty method 'template.get' |
			Add-Member -PassThru NoteProperty params @{output="extend"} |
			Add-Member -PassThru NoteProperty auth $session.result |
			Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		}
		$getJSON = getJSON $session.zabbixApiURL $objProxy
		if (!$getJSON.result){
			Write-Verbose "[get-ZabbixProxy] : Unable to retrieve host group $proxyName, aborting"
			return $false
		}
		else {return $getJSON }
				
	}
}

function get-ZabbixItem { 
	<#
	.Synopsis
		Retrieves items, you can search for host/hostgroup and item description
	.Example
		get-ZabbixItem 
		get-zabbixItem -short
		get-zabbixItem -itemDescription "CPU Load" -hostName "host1.lan"
		get-zabbixItem -itemDesciption "CPU Load" -hostGroupName "Group 1"
		get-zabbixItem -hostGroupName "Group 1" -hostPattern "apache"
	.Parameter itemDescription
		provide an item description pattern to limit the scope of the search, not mandatory
	.Parameter hostName
		provide a host name to limit the scope of the search, not compatible with -hostGroupName and -hostPattern, not mandatory
	.Parameter hostID
		provide a host id to limit the scope of the search, not compatible with -hostGroupName nor -hostName, not mandatory
	.Parameter hostGroupName
		provide a host group name to limit the scope of the search, not compatible with -hostName, not mandatory
	.Parameter hostPattern
		provide a host name pattern to limit the scope of the search to that patter, not compatible with -hostName, not mandatory
	.Parameter short
		short version, returns only itemids. Usefull because JSON input is limited to 2MB in size... :( 
	
	.Notes
		NAME: get-ZabbixItem
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$False,
		HelpMessage="provide an item description pattern to limit the scope of the search")]
		[String] $itemDescription
	,
		[Parameter(ParameterSetName="hostname",
		Mandatory=$False,
		HelpMessage="provide a host name to limit the scope of the search, not compatible with -hostGroupName and -hostPattern")]
		[String] $hostName
	,
		[Parameter(ParameterSetName="hostid",
		Mandatory=$False,
		HelpMessage="provide a host id to limit the scope of the search, not compatible with -hostGroupName nor -hostName")]
		[String] $hostId
	,
		[Parameter(ParameterSetName="hostgroup",
		Mandatory=$False,
		HelpMessage="provide a host group name to limit the scope of the search, not compatible with -hostName")]
		[String] $hostGroupName
	,
		[Parameter(ParameterSetName="hostgroup",
		Mandatory=$False,
		HelpMessage="provide a host name pattern to limit the scope of the search to that patter, not compatible with -hostName")]
		[String] $hostPattern
	,
		[Parameter(Mandatory=$False,
		HelpMessage="short version, returns only itemids. Usefull because JSON input is limited to 2MB in size... :( ")]
		[Switch] $short
	)
	Process{
		
		#construct the params
		$params=@{}
		
		#contruct the "hostids" param. 
		#Fist : let's check if we have some parameters that imply a second request
			#Because "search group and filter on host name" technique doesnt seem to work, we provide a list of host ids
		if ( ($hostGroupName) -and ($hostPattern) ) {
			$hosts=get-ZabbixHost -hostGroupName $hostGroupName -hostPattern $hostPattern -short
			$query=1
		}
		elseif ($hostGroupName) {
			$hosts=get-ZabbixHost -hostGroupName $hostGroupName -short
			$query=1
		}
		elseif ($hostPattern) {
			$hosts=get-ZabbixHost -hostPattern $hostPattern -short
			$query=1
		}
		#if any host filtering is asked and a proper list of host has been returned by get-zabbixhost, format the "hostids" param
		if ( ($query -eq 1) -and ($hosts) ) {
			$hostids=@()
			foreach ($hostid in $hosts.result)
				{
					$hostids += $hostid.hostid
				}
			$params.add("hostids",$hostids)
		}
		#Then let's look for some parameters that do not imply a second request (-hostId or -hostName)
		if ($hostId) {
			$params.add("hostids",$hostId)
		}
		elseif ($hostName) {$params.add("host",$hostName)}
		
		#construct the "search description" param
		if ($itemDescription)  {	
				$search=@{}
				$search.add("description", $itemDescription) 
				$params.add("search",$search)
		}
		#finish the params
		if ($short) {$params.add("output", "shorten")}
		else {$params.add("output", "extend") }
		
		#construct the JSON object	
		$objitem = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'item.get' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json
		
		#return $objitem
		
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objitem -errorMsg "[get-ZabbixItem] : Unable to retrieve items, aborting"
		#>
	}
}

function create-zabbixGraph {
	<#
	.Synopsis
		Create a graph in Zabbix with items matching a description, with hosts from a host group , optionally you can specify a host pattern. Very specific I know..;but it does the trick for me 
	.Example
		 create-zabbixGraph -graphName "Linux servers CPU Load" -itemDescription "CPU Load" -hostGroupName "Linux servers" 
		 create-zabbixGraph -graphName "Apache servers CPU Load" -itemDescription "CPU Load" -hostGroupName "Liunx servers" -hostPattern "Apache"
	.Parameter graphName
		provide a name for the Grpah, mandatory
	.Parameter hostGroupName
		provide a host group name to limit the scope of the search
	.Parameter itemDescription
		provide an item description pattern to limit the scope of the search, mandatory
	.Parameter hostPattern
		provide a host name pattern to limit the scope of the search to that pattern
	.Notes
		NAME: create-ZabbixGraph
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$True,
		HelpMessage="provide a name for the Grpah")]
		[String] $graphName
	,
		[Parameter(Mandatory=$True,
		HelpMessage="provide an item description pattern to limit the scope of the search")]
		[String] $itemDescription
	,
		[Parameter(ParameterSetName="hostgroup",
		Mandatory=$True,
		HelpMessage="provide a host group name to limit the scope of the search")]
		[String] $hostGroupName
	,
		[Parameter(ParameterSetName="hostgroup",
		Mandatory=$False,
		HelpMessage="provide a host name pattern to limit the scope of the search to that pattern")]
		[String] $hostPattern
		
	)
	Process {
		
		#create a list of colours to be used in the graph
		$colours=@("009900", "000099", "666666", "990000", "009999", "990099", "00EE00",
		"3333FF", "FF3333", "EE00EE", "FFFF33", "CCCCCC", "000066", "C04000", "800000", 
		"191970", "3EB489", "FFDB58", "000080", "CC7722","808000", "FF7F00", "002147", 
		"AEC6CF", "836953", "CFCFC4", "77DD77", "F49AC2","FFB347", "FFD1DC", "B39EB5",
		"FF6961", "CB99C9", "FDFD96", "FFE5B4", "D1E231","8E4585", "FF5A36", "701C1C",
		"FF7518", "69359C", "E30B5D", "826644", "FF0000","414833", "65000B", "002366",
		"E0115F", "B7410E", "FF6700", "F4C430", "FF8C69","C2B280", "967117", "ECD540", 
		"082567" )
		
		#SO, in the end we want a graph with a list of items sorted by host name...no easy way to do this...
		#1st, get a sorted list of hosts
		if ($hostPattern) {
			$hosts = get-ZabbixHost -hostGroupName $hostGroupName -hostPattern $hostPattern -short -sort
		}
		else {
			$hosts = get-ZabbixHost -hostGroupName $hostGroupName -short -sort
		}
		if (!$hosts) {
			Write-Error "No hosts returned, aborting"
			return $false
		}
		#then loop through the hosts, get the item(s) and add them to the list
		$params=@{}
		$gitems=@() 
		$c=0 #index for the $colours array
		foreach ($h in $hosts.result) {
			$items = get-ZabbixItem -hostId $h.hostid -itemDescription $itemDescription -short
			if ($items.result) {
				foreach ($itemid in $items.result) {
					$gitem=@{}
					$gitem.add("itemid", $itemid.itemid)
					$gitem.add("color", $colours[$c])
					$gitem.add("yaxisside", "0")
					$gitems += $gitem
					$c += 1
				}
			}
			else {Write-debug "[create-zabbixGraph] : no items matching $itemdescription, continuing"}
		
		}
		$params.add("gitems", $gitems)
		$params.add("name", $graphName)
		$params.add("width", "900")
		$params.add("height", "200")
		
		
		#construct the JSON object	
		$objgraph = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'graph.create' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json -depth 4
		
		#return $objgraph
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objgraph -errorMsg "[create-zabbixgraph] : Unable to create graph, aborting"
	}

}

function get-ZabbixGraphItem { 
	<#
	.Synopsis
		Retrieves Zabbix graph items
	.Example
		get-zabbixGraphItem -graphID 123456 
		get-zabbixGraphItem -graphID 123456 -short -sort
	.Parameter graphId
		Provide a single graph id to limit the scope of the search, mandatory
	.Parameter short
		Returns only item ids, not mandatory
	.Parameter sort
		sort by host name, not mandatory
	.Notes
		NAME: get-ZabbixGraphItem
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(ParameterSetName="graphid",
		Mandatory=$True,
		HelpMessage="provide a single graph id to limit the scope of the search, mandatory")]
		[String]$graphId
		,
		[Parameter(Mandatory=$False,
		HelpMessage="short version, returns only item ids. Usefull because JSON input is limited to 2MB in size... :( ")]
		[Switch] $short	
		,
		[Parameter(Mandatory=$False,
		HelpMessage="sort by host name")]
		[Switch] $sort
	)
	Process {
		#construct the params
		$params=@{}
		$params.add("graphids", $graphId)
		if ($short) {$params.add("output", "shorten")}
		else {$params.add("output", "extend") }
		
		if ($sort) {$params.add("sortfield","gitemid") }
		
		#construct the JSON object
		$objGraphItem = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'graphitem.get' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	

		#make the request and evaluate it
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objGraphItem -errorMsg "[get-zabbixGraphItem] : Unable to retrieve hosts, aborting." -noDataMsg "[get-zabbixGraphItem
		] : No host matching the description"
	}
}

function get-ZabbixGraph { 
	<#
	.Synopsis
		Retrieves Zabbix graphs, deprecated, you should use get-zabbixGrpahByHost, get-ZabbixGraphByID, get-ZabbixGraphByGroup instead
	.Example
		get-zabbixGraph -graphID 123456
		get-zabbixGraph -hostID 123456
		get-zabbixGraph -hostName "myhost.lan"
		get-zabbixGrpah -graphName "CPU Load" -hostName "apache01.lan"
	.Parameter graphId
		provide a single graph id to limit the scope of the search, not compatible with -hostId nor -hostName
	.Parameter graphName
		provide a single graph name to limit the scope of the search, not mandatory
	.Parameter hostID
		provide a single host id to limit the scope of the search, not compatible with -graphId nor -hostName
	.Parameter hostName
		provide a single host name to limit the scope of the search, not compatible with -hostId nor -graphId
	.Parameter short
		Returns only item ids
	.Notes
		NAME: get-ZabbixGraph
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(ParameterSetName="retrieveFromGraphId",
		Mandatory=$False,
		HelpMessage="provide a single graph id to limit the scope of the search, not compatible with -hostId nor -hostName")]
		[String]$graphId
		,
		[Parameter(ParameterSetName="filterByHostId",
		Mandatory=$False,
		HelpMessage="provide a single host id to limit the scope of the search, not compatible with -graphId nor -hostName")]
		[String]$hostId
		,
		[Parameter(ParameterSetName="filterByHostName",
		Mandatory=$False,
		HelpMessage="provide a single host name to limit the scope of the search, not compatible with -hostId nor -graphId")]
		[String]$hostName
		,
		[Parameter(Mandatory=$False,
		HelpMessage="provide a single graph name to limit the scope of the search, not mandatory")]
		[String]$graphName
		,
		[Parameter(Mandatory=$False,
		HelpMessage="short version, returns only item ids. Usefull because JSON input is limited to 2MB in size... :( ")]
		[Switch] $short		
	)
	Process {
		#construct the params
		$params=@{}
		if ($grapId) { $params.add("graphids", $graphId) }
		elseif ($groupId) { $params.add("groupids", $groupId) }
		elseif ($hostId) { $params.add("hostids", $hostId) }
		elseif ($hostName) {
			$h=get-ZabbixHost -hostName $hostName
			if ($h) {
				$params.add("hostids",$h.result.hostid)
			}
			else { Write-Verbose "[get-ZabbixGraph] : no host name matching, ignoring this parameter" }
		}
		elseif ($graphName) { 
			$filter=@{}
			$filter= @{"name" = "$graphName"}
			$params.add("filter",$filter)
		}
		if ($short) {$params.add("output", "shorten")}
		else {$params.add("output", "extend") }
		
		#construct the JSON object
		$objGraphItem = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'graph.get' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	

		#return $objGraphItem
		
		#make the request and evaluate it
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objGraphItem -errorMsg "[get-zabbixGraph] : Unable to retrieve graph, aborting." -noDataMsg "[get-zabbixGraph] : No graph matching the description"
	}
}

function get-ZabbixGraphByID { 
	<#
	.Synopsis
		Retrieves a Zabbix graph from it's ID
	.Example
		get-zabbixGrpahByID -graphID 123456
	.Parameter graphID
		provide a single graph id to limit the scope of the search, mandatory
	.Notes
		NAME: get-ZabbixGraphByID
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(ParameterSetName="hostId",
		Mandatory=$True,
		HelpMessage="provide a single graph id to limit the scope of the search, mandatory")]
		[String]$graphId	
	)
	Process {
		#construct the params
		$params=@{}
		$params.add("graphids", $graphId)
		$params.add("output", "extend")
		
		#construct the JSON object
		$objGraphItem = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'graph.get' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	

		#return $objGraphItem
		
		#make the request and evaluate it
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objGraphItem -errorMsg "[get-zabbixGraphByID] : Unable to retrieve graph, aborting." -noDataMsg "[get-zabbixGraph] : No graph matching the description"
	}
}

function get-ZabbixGraphByHost { 
	<#
	.Synopsis
		Retrieves Zabbix graphs for a specific host
	.Example
		get-zabbixGraphByHost -hostID 123456
		get-zabbixGraphByHost -hostID 123456 -graphDescription "CPU"
		get-zabbixGraphByHost -hostName "myhost.lan" 
		get-zabbixGraphByHost -hostName "myhost.lan" -graphDescription "CPU"
		get-zabbixGraphByHost -hostName "myhost.lan" -short
	.Parameter hostID
		provide a single host id to limit the scope of the search, not compatible with -hostName, mandatory
	.Parameter hostName
		provide a single host name to limit the scope of the search, not compatible with -hostId, mandatory
	.Parameter graphDescription
		provide a single graph name to limit the scope of the search. If the description has no match, Zabbix API will ignore it and return all graphs! not mandatory
	.Parameter short
		Returns only item ids
	.Notes
		NAME: get-ZabbixGraphByHost
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(ParameterSetName="hostId",
		Mandatory=$True,
		HelpMessage="provide a single host id to limit the scope of the search, not compatible with -hostName")]
		[String]$hostId
		,
		[Parameter(ParameterSetName="hostName",
		Mandatory=$True,
		HelpMessage="provide a single host name to limit the scope of the search, not compatible with -hostId")]
		[String]$hostName
		,
		[Parameter(Mandatory=$False,
		HelpMessage="provide a single graph name to limit the scope of the search, not mandatory")]
		[String]$graphDescription
		,
		[Parameter(Mandatory=$False,
		HelpMessage="short version, returns only item ids. Usefull because JSON input is limited to 2MB in size... :( ")]
		[Switch] $short		
	)
	Process {
		#construct the params
		$params=@{}
		
		#depending on the parameter set we might need an additional request to get the host id from the name
		if ($hostId) 
			{ $params.add("hostids", $hostId) }
		elseif ($hostName) {
			$h=get-ZabbixHost -hostName $hostName
			if ($h.result.Count -eq 1) {
				$params.add("hostids",$h.result.hostid)
			}
			else { 
				Write-Verbose "[get-ZabbixGraph] : no hostname matching, aborting" 
				return $false 
			}
		}
		
		#if it is required, try to limit the search with a description
		if ($graphDescription) { 
			$search=@{}
			$search= @{"name" = "$graphDescription"}
			$params.add("search",$search)
		}
		
		if ($short) {$params.add("output", "shorten")}
		else {$params.add("output", "extend") }
		
		#construct the JSON object
		$objGraphItem = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'graph.get' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	

		#return $objGraphItem
		
		#make the request and evaluate it
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objGraphItem -errorMsg "[get-zabbixGraphByHost] : Unable to retrieve graph, aborting." -noDataMsg "[get-zabbixGraphByHost] : No graph matching the description"
	}
}

function get-ZabbixGraphByHostGroup { 
	<#
	.Synopsis
		Retrieves Zabbix graphs for a specific group
	.Example
	.Parameter hostGroupID
		provide a single group id to limit the scope of the search, not compatible with -groupName, mandatory
	.Parameter hostGroupName
		provide a single group name to limit the scope of the search, not compatble with -groupID, Id mandatory
	.Parameter graphDescription
		provide a single graph name to limit the scope of the search. If the description has no match, Zabbix API will ignore it and return all graphs! not mandatory
	.Parameter short
		Returns only item ids
	.Notes
		NAME: get-ZabbixGraphByHostGroup
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(ParameterSetName="groupId",
		Mandatory=$True,
		HelpMessage="provide a single group id to limit the scope of the search, not compatible with -groupName")]
		[String]$hostGroupId
		,
		[Parameter(ParameterSetName="groupName",
		Mandatory=$True,
		HelpMessage="provide a single group name to limit the scope of the search, not compatible with -groupId")]
		[String]$hostGroupName
		,
		[Parameter(Mandatory=$False,
		HelpMessage="provide a single graph name to limit the scope of the search, not mandatory")]
		[String]$graphDescription
		,
		[Parameter(Mandatory=$False,
		HelpMessage="short version, returns only item ids. Usefull because JSON input is limited to 2MB in size... :( ")]
		[Switch] $short		
	)
	Process {
		#construct the params
		$params=@{}
		
		#depending on the parameter set we might need an additional request to get the group id from the name
		if ($hostGroupId) 
			{ $params.add("groupids", $hostGroupId) }
		elseif ($hostGroupName) {
			$h=get-ZabbixHostGroup -hostGroupName $hostGroupName
			if ($h.result.Count -eq 1) {
				$params.add("groupids",$h.result.groupid)
			}
			else { 
				Write-Verbose "[get-ZabbixGraph] : no hostgroup name matching, aborting" 
				return $false 
			}
		}
		
		#if it is required, try to limit the search with a description
		if ($graphDescription) { 
			$search=@{}
			$search= @{"name" = "$graphDescription"}
			$params.add("search",$search)
		}
		
		if ($short) {$params.add("output", "shorten")}
		else {$params.add("output", "extend") }
		
		#construct the JSON object
		$objGraphItem = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'graph.get' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	

		#return $objGraphItem
		
		#make the request and evaluate it
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objGraphItem -errorMsg "[get-zabbixGraphByGroup] : Unable to retrieve graph, aborting." -noDataMsg "[get-zabbixGraphByGroup] : No graph matching the description"
	}
}

function download-ZabbixGraphImage {
	<#
	.Synopsis
		This function downloads a Zabbix Graph Image to a PNG file
	.Example
		download-zabbixGraphImage -graphID 123456 -zabbixWebSession $webSession -outFile ".image.png" -startDate 20131012011315 -endate 20131013011315
	.Parameter graphId
		provide a Zabbix Graph ID, mandatory
	.Parameter zabbixWebSession
		provide a websession object with a valid cookie for the graph URL.You will probably have to set it with setcookies. mandatory
	.Parameter outFile
		Enter a path and file name. If you omit the path, the default is the current location, mandatory
	.Parameter startDate
		Starting date as YYYYMMDDhhmmss, mandatory
	.Parameter endDate
		Ending date as YYYYMMDDhhmmss, should not be greater than current date, mandatory
	.Notes
		NAME: download-ZabbixGraphImage
		AUTHOR: Simon Morand (MBVSI)
	#>
[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$True
		,HelpMessage="provide a Zabbix Graph ID, mandatory")]
		[String]$graphId
	,
		[Parameter(Mandatory=$True
		,HelpMessage="provide a websession object with a valid cookie for the graph URL.You will probably have to set it with setcookies. mandatory")]
		[Microsoft.PowerShell.Commands.WebRequestSession]$zabbixWebSession
	,
		[Parameter(Mandatory=$True
		,HelpMessage="Enter a path and file name. If you omit the path, the default is the current location, mandatory")]
		[string]$outFile
	,
		[Parameter(Mandatory=$True
		,HelpMessage="Starting date as YYYYMMDDhhmmss, mandatory")]
		[string]$startDate
	,
		[Parameter(Mandatory=$True
		,HelpMessage="Ending date as YYYYMMDDhhmmss, should not be greater than current date, mandatory")]
		[string]$endDate
	)
	
	Process {
		#get the chart2.php page from the $session info
		$zabbixGraphUrl=$session.$zabbixApiURL.trimend("api_jsonrpc.php")
		$zabbixGraphUrl+="chart2.php"
	
		#let's check that the websession has a valid cookie
		if ($zabbixWebSession.Cookies.GetCookies($zabbixGraphUrl).Count -ne 1) {
			Write-InRed "No valid cookie for $zabbixGraphUrl found in zabbixWebSession, aborting"
			return $false
		}
		
		#let's convert our string parameters into dates 
		try {
			$sd = [datetime]::ParseExact($startDate,"yyyyMMddHHmmss",$null)
			$ed = [datetime]::ParseExact($endDate,"yyyyMMddHHmmss",$null)
		}
		catch [system.datetime]{
			Write-InRed "Problem with a date paramenter, error message is $Error"
			return $false
		}
		
		#let's check that -endDate is not after current date or before -startDate
		if ($ed -gt (Get-Date)) {
			Write-InRed "Parameter -endDate is after current date, aborting"
			return $false
		}
		elseif ($ed -lt $sd) {
			Write-InRed "Parameter -endDate is before -startDate, aborting"
			return $false
		}
		#Now that we know the dates make sense, let's get the difference in seconds between the 2
		[string]$diffSec = ($ed-$sd).TotalSeconds
		
		#create the request 
		[string]$params
		$params+="?"
		$params += "graphid=$graphId"
		$params += "&stime=$startDate"
		$params += "&period=$diffSec"
		
		
		$graph = Invoke-WebRequest -Uri ($zabbixGraphUrl+$params) -WebSession $zabbixWebSession -MaximumRedirection 3 -OutFile $outFile -PassThru
		if ($graph.StatusDescription -ne "OK") {
			Write-InRed "Something went wrong with the request, status code is $graph.statusCode"
			return $false
		}
		#check if the zabbix error picture has been returned, easy, it's 446B long
		if ($graph.RawContentLength -eq 446) {
			Write-Verbose "[download-ZabbixGraphImage] : Error downloading image for graph $graphid, aborting"
			return $false
		
		}
		else {
			return $True
		}

	}

}

function download-ZabbixLastDataImage {
	<#
	.Synopsis
		This function downloads a Zabbix Last Data Image to a PNG file
	.Example
		download-zabbixLastDataImage -itemID 123456 -zabbixWebSession $webSession -outFile ".image.png" -startDate 20131012011315 -endate 20131013011315
	.Parameter itemID
		provide a Zabbix item ID, mandatory
	.Parameter zabbixWebSession
		provide a websession object with a valid cookie for the graph URL.You will probably have to set it with setcookies. mandatory
	.Parameter outFile
		Enter a path and file name. If you omit the path, the default is the current location, mandatory
	.Parameter startDate
		Starting date as YYYYMMDDhhmmss, mandatory
	.Parameter endDate
		Ending date as YYYYMMDDhhmmss, should not be greater than current date, mandatory
	.Notes
		NAME: download-ZabbixLastDataImage
		AUTHOR: Simon Morand (MBVSI)
	#>
[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$True
		,HelpMessage="provide a Zabbix item ID, mandatory")]
		[String]$itemID
	,
		[Parameter(Mandatory=$True
		,HelpMessage="provide a websession object with a valid cookie for the graph URL.You will probably have to set it with setcookies. mandatory")]
		[Microsoft.PowerShell.Commands.WebRequestSession]$zabbixWebSession
	,
		[Parameter(Mandatory=$True
		,HelpMessage="Enter a path and file name. If you omit the path, the default is the current location, mandatory")]
		[string]$outFile
	,
		[Parameter(Mandatory=$True
		,HelpMessage="Starting date as YYYYMMDDhhmmss, mandatory")]
		[string]$startDate
	,
		[Parameter(Mandatory=$True
		,HelpMessage="Ending date as YYYYMMDDhhmmss, should not be greater than current date, mandatory")]
		[string]$endDate
	)
	
	Process {
		#get the chart.php page from the $session info
		$zabbixLastDataUrl=$session.$zabbixApiURL.trimend("api_jsonrpc.php")
		$zabbixLastDataUrl+="chart.php"
	
		#let's check that the websession has a valid cookie
		if ($zabbixWebSession.Cookies.GetCookies($zabbixLastDataUrl).Count -ne 1) {
			Write-InRed "No valid cookie for $zabbixLastDataUrl found in zabbixWebSession, aborting"
			return $false
		}
		
		#let's convert our string parameters into dates 
		try {
			$sd = [datetime]::ParseExact($startDate,"yyyyMMddHHmmss",$null)
			$ed = [datetime]::ParseExact($endDate,"yyyyMMddHHmmss",$null)
		}
		catch [system.datetime]{
			Write-InRed "Problem with a date paramenter, error message is $Error"
			return $false
		}
		
		#let's check that -endDate is not after current date or before -startDate
		if ($ed -gt (Get-Date)) {
			Write-InRed "Parameter -endDate is after current date, aborting"
			return $false
		}
		elseif ($ed -lt $sd) {
			Write-InRed "Parameter -endDate is before -startDate, aborting"
			return $false
		}
		#Now that we know the dates make sense, let's get the difference in seconds between the 2
		[string]$diffSec = ($ed-$sd).TotalSeconds
		
		#create the request 
		[string]$params
		$params+="?"
		$params += "itemid=$itemId"
		$params += "&stime=$startDate"
		$params += "&period=$diffSec"
		
		
		$lastData = Invoke-WebRequest -Uri ($zabbixLastDataUrl+$params) -WebSession $zabbixWebSession -MaximumRedirection 3 -OutFile $outFile -PassThru
		if ($lastData.StatusDescription -ne "OK") {
			Write-InRed "Something went wrong with the request, status code is $lastData.statusCode"
			return $false
		}
		#check if the zabbix error picture has been returned, easy, it's 446B long
		if ($lastData.RawContentLength -eq 446) {
			Write-Verbose "[download-ZabbixlastDataImage] : Error downloading image for item $itemid, aborting"
			return $false
		
		}
		else {
			return $True
		}

	}

}

function create-zabbixGraphFromGitems {
	<#
	.Synopsis
		pass a hash table with itemid->color entries to create a graph with
	.Example
		 create-zabbixGraphFromGitems -graphName "My new graph" -gitems $arrayOfGitems
	.Parameter graphName
		provide a name for the Graph
	.Parameter gitems
		provide an array filled with itemid+color hash tables entries
	.Notes
		NAME: create-zabbixGraphFromGitems
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$True,
		HelpMessage="provide a name for the Graph")]
		[String] $graphName
	,
		[Parameter(Mandatory=$True,
		HelpMessage="provide an array filled with itemid+color hash tables entries")]
		[Array] $gitems
	)
	Process {
		
		$params=@{}
		$params.add("gitems", $gitems)
		$params.add("name", $graphName)
		$params.add("width", "900")
		$params.add("height", "200")
		
		
		#construct the JSON object	
		$objgraph = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'graph.create' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json -depth 4
		
		#return $objgraph
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objgraph -errorMsg "[create-zabbixgraph] : Unable to create graph, aborting"
	}

}

function clone-zabbixGraph {
<#
	.Synopsis
		Copy a graph and all it's elements into a new one with a different name. You can choose to delete the original
	.Example
		 
	.Parameter 
		
	.Notes
		NAME: clone-ZabbixGraph
		AUTHOR: Simon Morand (MBVSI)
	#>
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$True,
		HelpMessage="id of the original graph")]
		[String] $origGraphid
		,
		[Parameter(Mandatory=$True,
		HelpMessage="name for the new graph")]
		[String] $newGraphName
		,
		[Parameter(Mandatory=$false,
		HelpMessage="delete the original graph if true")]
		[switch] $delete
	)
	
	Process {
		
		#let's retrieve all items from the original graph
		$originGraphItems=get-ZabbixGraphItem -graphId $origGraphId -sort
		if (!$originGraphItems.result) {
			Write-Verbose "[clone-zabbixGraph] : could not retrieve items from original graph, aborting"
			return $false
		}
		
		#now we can build a gitems array, as needed in the graph.create method
		$newGraphItems=@()
		foreach ($line in $originGraphItems.result) {
			$gitem=@{}
			$gitem.add("itemid", $line.itemid)
			$gitem.add("color", $line.color)
			$gitem.add("yaxisside", "0")
			$newGraphItems += $gitem
		}
		#return $newGraphItems
		create-zabbixGraphFromGitems -gitems $newGraphItems -graphName $newGraphName
		
	}

}

function get-ZabbixHostByID {
<#
	.Synopsis
		Rerieves a host from it's ID
	.Example
		get-zabbixHostByID -hostID 123456
		get-zabbixHostByID -hostID 123456 -selectGroups
		get-zabbixHostByID -hostID 123456 -selectGraphs
	.Parameter hostId
		provide a host ID, mandatory
	.Parameter selectGroups
		Return the host groups that the host belongs to in the groups property, not mandatory
	.Parameter selectGraphs
		Return the graphs attached to the host in the graphs property, not mandatory
	.Notes
		NAME: get-ZabbixHostByID
		AUTHOR: Simon Morand (MBVSI)
	#>
	
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$True,
		HelpMessage="provide a host ID, mandatory")]
		[String]$hostID
		,
		[Parameter(Mandatory=$False,
		HelpMessage="Return the host groups that the host belongs to in the groups property, not mandatory")]
		[Switch] $selectGroups
		,
		[Parameter(Mandatory=$False,
		HelpMessage="Return the graphs attached to the host in the graphs property, not mandatory")]
		[Switch] $selectGraphs
	)
	Process {
		
		$params=@{}
	
		$params.add("hostids", $hostID)
		
		if ($selectGroups) {$params.add("select_groups","refer")}
		
		if ($selectGraphs) {$params.add("select_graphs","refer")}
		
		if ($short) {$params.add("output", "shorten")}
		else {$params.add("output", "extend") }
		
		#construct the JSON object
		$objHost = (New-Object PSObject | Add-Member -PassThru NoteProperty jsonrpc '2.0' |
		Add-Member -PassThru NoteProperty method 'host.get' |
		Add-Member -PassThru NoteProperty params $params |
		Add-Member -PassThru NoteProperty auth $session.result |
		Add-Member -PassThru NoteProperty id '2') | ConvertTo-Json 	
		
		#return $objHost
		
		#make the request and evaluate it
		evaluate-JSON -jsonApiUrl $session.zabbixApiURL -jsonObj $objHost -errorMsg "[get-ZabbixHostByID] : Unable to retrieve host, aborting." -noDataMsg "[get-ZabbixHostByID] : No host matching the description"
	
	
	}
}

function download-ZabbixImageFromCSV {
	<#
	.Synopsis
		This function downloads a list of Zabbix Last Data OR Graph Image from a CSV file to a PNG file
	.Paramter csvFile
		Enter a path to a csv file containing images to download , mandatory
	.Parameter outPath
		Enter a directory to place images in, mandatory
	.Parameter startDate
		Starting date of the graph, mandatory
	.Parameter endDate
		Ending date of the graph, mandatory
	.Parameter zabbixWebSession
		provide a websession object with a valid cookie for graph and lastdata URLs.You will probably have to set it with setcookies. mandatory
	.Notes
		NAME: download-ZabbixImageFromCSV
		AUTHOR: Simon Morand (MBVSI)
	#>

[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$True
		,HelpMessage="Enter a path to a csv file containing images to download , mandatory")]
		[ValidateScript({Test-Path $_ -PathType 'leaf'})]
		[string]$csvFile
	,
		[Parameter(Mandatory=$True
		,HelpMessage="Enter a directory to place images in, mandatory")]
		[ValidateScript({Test-Path $_ -PathType 'container'})]
		[string]$outPath
	,
		[Parameter(Mandatory=$True
		,HelpMessage="Starting date as YYYYMMDDhhmmss, mandatory")]
		[string]$startDate
	,
		[Parameter(Mandatory=$True
		,HelpMessage="Ending date as YYYYMMDDhhmmss, should not be greater than current date, mandatory")]
		[string]$endDate
	,	
		[Parameter(Mandatory=$True
		,HelpMessage="provide a websession object with a valid cookie for the graph URL, mandatory")]
		[Microsoft.PowerShell.Commands.WebRequestSession]$zabbixWebSession
	)
	Process {
		#need to connect-zabbix first

		#import csv file. file must be like this : type;hostname;description;filename where :
		#	type : can be "lastdata" or "graph"
		#	hostname : name of the zabbix host	
		#	description : can be an item description if type is "lastdata" or a graph description if type is "graph"
		#	filename : name of the png file to be created (without path!)
		#delimiter for this csv file must be ";"

		
		#First let's check that the two date parameters make sense
		try {
			$sd = [datetime]::ParseExact($startDate,"yyyyMMddHHmmss",$null)
			$ed = [datetime]::ParseExact($endDate,"yyyyMMddHHmmss",$null)
		}
		catch [system.datetime]{
			Write-InRed "Problem with a date paramenter, error message is $Error"
			return $false
		}
		
		#let's check that -endDate is not after current date or before -startDate
		if ($ed -gt (Get-Date)) {
			Write-InRed "Parameter -endDate is after current date, aborting"
			return $false
		}
		elseif ($ed -lt $sd) {
			Write-InRed "Parameter -endDate is before -startDate, aborting"
			return $false
		}
		
		#then import the csv file, we already check in parameters declaration that path exists
		$csv = Import-Csv -Path $csvFile -Delimiter ";"


		#loop through the file and download pictures
		[datetime]$processStartDate=Get-Date 
		[int]$nbrImages

		foreach ($line in $csv) {
			#if type is "graph" we call download-ZabbixGraphImage
			if($line.type -eq "graph") {
				$g=get-ZabbixGraphByHost -hostName $line.hostName -graphDescription $line.description
				if ($g.result.Count -eq 1) {
					$outfile = $outPath+$line.fileName
					if (download-ZAbbixGraphImage -graphID $g.result.graphid -zabbixWebSession $zabbixWebSession -outFile $outfile -startDate $startDate -endDate $endDate) {
						Write-host "graph " $line.description " for host " $line.hostName " downloaded successfully"
						$nbrImages ++
					}
					else {
						Write-host "A problem occured while downloading graph " $line.description "for host "$line.hostName" , skeeping"
					}
				}
				else {
					Write-host "Could not retrieve a single graph matching description" $line.description " for host "$line.hostName" , skeeping"
				}
			
			}
			elseif($line.type -eq "lastdata") {
				#if type is "lastdata" we call download-ZabbixLastDataImage
				$i=get-ZabbixItem -hostName $line.hostName -itemDescription $line.description
				if ($i.result.Count -eq 1) {
					$outfile = $outPath+$line.fileName
					if (download-ZabbixLastDataImage -itemID $i.result.itemid -zabbixWebSession $zabbixWebSession -outFile $outfile -startDate $startDate -endDate $endDate ) {
						Write-host "Last Data " $line.description " for host " $line.hostName " downloaded successfully"
						$nbrImages ++
		
					}
					else {
						Write-host "A problem occured while downloadding graph " $line.description "for host "$line.hostName" , skeeping"
					}
				}
				else {
					Write-host "Could not retrieve a single item matching description " $line.itemDescription " for host "$line.hostName", skeeping"
				}
			
			}
			elseif ($line.type -eq "comment") {
				Write-Host "`n" $line.hostName
			}
			else {
				Write-host "Type : "$line.type"  is not correct, must be [lastdata|graph|comment], skipping"
			}
		}
	}
}
