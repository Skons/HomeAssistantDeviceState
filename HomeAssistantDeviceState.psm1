$Script:CurrentState = $null

Function Get-DevicePDO {
	<#
	.SYNOPSIS
		Get the Physical Device Object name of a device
	.DESCRIPTION
		Get the Physical Device Object name of a device
	.PARAMETER DeviceClass
		Get all devices of the specified device class
	.PARAMETER FriendlyName
		The friendly name of the device
	.EXAMPLE
		Get-DevicePDO
	.EXAMPLE
		Get-DevicePDO -DeviceClass camera
	#>
	[CmdLetBinding(DefaultParameterSetName='DeviceClass')]
	param(
		[Parameter(Mandatory=$False,ParameterSetName="DeviceClass")]
		[ValidateSet('Camera','AudioEndpoint','Media','Image')]
		[string[]]$DeviceClass=@('Camera','Media','Image'),
		[Parameter(Mandatory=$True,ParameterSetName="FriendlyName")]
		[string]$FriendlyName
	)
	
	if ($DeviceClass) {
		foreach ($Class in $DeviceClass) {
			[array]$DeviceClassWMIStrings += "DeviceClass='$Class'"
		}
		[string]$Query = "select * from Win32_PnPSignedDriver where $($DeviceClassWMIStrings -join ' or ')"
	}
	else {
		[string]$Query = "select * from Win32_PnPSignedDriver where friendlyname='$FriendlyName'"
	}

	Write-Debug "WMI Query '$Query'"
	return Get-WmiObject -Query $Query | Select-Object FriendlyName,DeviceClass,PDO
}

Function Get-DeviceInUseByProcess {
	<#
	.SYNOPSIS
		Find process that has device in use
	.DESCRIPTION
		Find processes that use a specific device. This can be a camera or a microphone
	.PARAMETER Handle
		Path to Sysinternals handle.exe
	.PARAMETER ProcessName
		Provide one or more (partial) process names to search the device handle
	.PARAMETER PDO
		The Physical Device Object name of the device to search the handle for
	.EXAMPLE
		Get-DeviceInUseByProcess -Handle .\handle64.exe -PDO \device\00000059
	#>
	[CmdLetBinding()]
	param(
		[Parameter(Mandatory=$True)]
		[string]$Handle,
		[Parameter(Mandatory=$False)]
		[string[]]$ProcessName,
		[Parameter(Mandatory=$True)]
		[string]$PDO
	)
 
	$Objects = @()
	$HandleData = @()
	if ($ProcessName) {
		foreach ($ProcName in $ProcessName) {
			Write-Verbose "Getting handles for process '$ProcName'"
			$HandleData += . $Handle -p $ProcName -a -nobanner
		}
	}
	else {
		Write-Verbose "Getting all process handles"
		$HandleData += . $Handle -a -nobanner
	}
	$Regex = "(^\S+)(\W+)(pid:)(\W+)(\d+)(\W+).*"
	for ($i=0;$i-lt$HandleData.Count;$i++) {
		if ($HandleData[$i] -like '----------*') {

		}
		elseif (![string]::IsNullOrWhiteSpace($HandleData[$i])) {
			if ($HandleData[$i] -match $Regex) {
				 $Object = '' | Select-Object ProcessName,ProcessId,PDO
				 $Object.ProcessName = $Matches[1]
				 $Object.ProcessId = $Matches[5]
				 Write-Debug "Found process '$($Object.ProcessName)' with PID '$($Object.ProcessId)'"
				 for ($j=$i+1;$j-lt$HandleData.Count;$j++) {
					if ($HandleData[$j] -match $Regex) {
						Write-Debug "Hit next process with name '$($Matches[1])'"
						$i=$j-1
						break
					}
					elseif ($HandleData[$j] -like "*$PDO") {
						Write-Verbose "Found Physical Device Object name '$PDO' within process '$($Object.ProcessName)'"
						$Object.PDO = $HandleData[$j]
					}
				 }
				 if ($Object.PDO) {
					$Objects += $Object
				 }
			}
		}
	}
	return $Objects
}

Function Set-HAEntityStateByDeviceInUse {
	<#
	.SYNOPSIS
		Change the state of an entity in HA based on device use
	.DESCRIPTION
		Set the state of an entity in HA based on if a device is in use
	.PARAMETER Handle
		See Get-DeviceInUseByProcess
	.PARAMETER ProcessName
		See Get-DeviceInUseByProcess
	.PARAMETER PDO
		See Get-DeviceInUseByProcess
	.PARAMETER Uri
		Uri to your homeassistant instance
	.PARAMETER Entity
		The entity to set a value
	.PARAMETER FoundState
		The value that will be set if the device is in use by a process
	.PARAMETER NotFoundState
		The value that will be set if the device is not in use
	.PARAMETER VaultName
		The name of the vault used with Microsoft.PowerShell.SecretsManagement. If not specified the default vault will be used
	.PARAMETER SecretName
		The name of the secret in your vault that has the HA Long Lived Access Token
	.PARAMETER Force
		If you want to force check the current state at HA, instead of relying on the state in memory of this module, specify this parameter
	.NOTES
		Install-Module -Name Microsoft.PowerShell.SecretsManagement -RequiredVersion 0.2.0-alpha1 -AllowPrerelease
	.EXAMPLE
		Set-HAEntityStateByDeviceInUse -ProcessName svchost -Handle .\handle64.exe -PDO "\Device\00000059" -Uri "http://hassio.local:8123/" -Entity 'input_boolean.bto_camera' -FoundStateValue 'On' -NotFoundStateValue 'Off' -SecretName HAToken -verbose
	.EXAMPLE
		Set-HAEntityStateByDeviceInUse -ProcessName svchost,HdxTeams -Handle .\handle64.exe -PDO "\Device\00000059" -Uri "http://hassio.local:8123/" -Entity 'input_boolean.bto_camera' -FoundStateValue 'On' -NotFoundStateValue 'Off' -SecretName HAToken -verbose
	#>
	[CmdLetBinding()]
	param(
		[Parameter(Mandatory=$True)]
		[string]$Handle,
		[Parameter(Mandatory=$False)]
		[string[]]$ProcessName,
		[Parameter(Mandatory=$True)]
		[string]$PDO,
		[Parameter(Mandatory=$True)]
		[uri]$Uri,
		[Parameter(Mandatory=$True)]
		[string]$Entity,
		[Parameter(Mandatory=$True)]
		[string]$FoundStateValue,
		[Parameter(Mandatory=$True)]
		[string]$NotFoundStateValue,
		[Parameter(Mandatory=$False)]
		[string]$VaultName,
		[Parameter(Mandatory=$True)]
		[string]$SecretName,
		[Parameter(Mandatory=$False)]
		[switch]$Force
	)

	#Get the secret
	$SecretParams = @{}
	if ($VaultName) {
		$SecretParams['Vault'] = $VaultName
	}
	[string]$LongLivedAccessToken = get-secret -Name $SecretName -AsPlainText @SecretParams -erroraction stop
	Write-Debug "Got token '$($LongLivedAccessToken)'"

	$endpoint = "$($uri.AbsoluteUri)api/states/$Entity"
	write-debug "Endpoint will be '$endpoint'"

	#Set parameters for Get-DeviceInUseByProcess
	$DeviceInUseByProcessParams = @{}
	if ($ProcessName) {
		$DeviceInUseByProcessParams['ProcessName'] = $ProcessName
	}

	$DeviceInUse = Get-DeviceInUseByProcess -Handle $Handle -PDO $PDO @DeviceInUseByProcessParams
	if ($DeviceInUse) {
		Write-Verbose "The device '$PDO' is in use, the state of '$Entity' should be '$FoundStateValue'"
		$StateToSend = $FoundStateValue
	}
	else {
		Write-Verbose "The device '$PDO' is not in use, the state of '$Entity' should be '$NotFoundStateValue'"
		$StateToSend = $NotFoundStateValue
	}

	$headers = @{Authorization = "Bearer $LongLivedAccessToken"}

	#Get the current state from HA if it is not known
	if ($null -eq $Script:CurrentState -or $Force) { #Get the entity state from HA
		$CurrentEntityState = Invoke-RestMethod -Uri $endpoint -Method 'get' -Headers $headers -UseBasicParsing -verbose:$False
		if ($CurrentEntityState.State -eq $StateToSend) {
			Write-Verbose "The current state '$($CurrentEntityState.State)' of '$Entity' at '$endpoint' is not changed"
			$Script:CurrentState = $CurrentEntityState.State
		}
	}

	#The current state of the entity is not correct, setting that state
	if ($Script:CurrentState -ne $StateToSend) {
		$Body = @"
{
	"state": "$StateToSend"
}
"@
		Write-Verbose "The current state '$Script:CurrentState' of '$Entity' is changed, setting it at '$endpoint' to '$StateToSend'"

		$response = Invoke-RestMethod -body $body -ContentType "application/json" -Uri $endpoint -Method 'post' -Headers $headers -UseBasicParsing -verbose:$False
		if ($response.State -eq $StateToSend) {
			write-verbose "Setting the state of '$Entity' was successfull"
			$Script:CurrentState = $StateToSend
		}
		return $Response
	}
}
