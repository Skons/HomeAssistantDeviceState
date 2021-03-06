$Script:CurrentState = $null

Function Get-ContentStoreDeviceUsedBy {
	<#
	.SYNOPSIS
		List the executables and/or Store apps that used a device
	.DESCRIPTION
		List the executables and/or Store apps that used a device. If the type is an executable, the path is also provided
	.EXAMPLE
		Get-ContentStoreDeviceUsedBy
	#>
	[CmdLetBinding()]
	Param()

	$TemplateObject = '' | Select-Object Type,Path,Name,Device,LastUsedStartTime,LastUsedStopTime
	$RegistryRoot = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore'
	Write-Verbose "Getting devices from '$RegistryRoot'"
	$Items = Get-ItemProperty $(join-path $RegistryRoot '\*\NonPackaged\*')
	$Items | foreach-Object {
		$Path = $_.PSChildName -replace '#','\'
		$Object = $TemplateObject.PSObject.Copy()
		$Object.Type = 'Executable'
		$Object.Path = $Path | split-path
		$Object.Name = $Path | split-path -leaf
		$Object.Device = $_.PSParentPath | Split-Path | Split-Path -Leaf
		$Object.LastUsedStartTime = [datetime]::FromFileTime($_.LastUsedTimeStart)
		$Object.LastUsedStopTime = 0
		if ($_.LastUsedTimeStop -ne 0) {
			$Object.LastUsedStopTime = [datetime]::FromFileTime($_.LastUsedTimeStop)
		}
		$Object
	}

	$Items = Get-ItemProperty $(join-path $RegistryRoot '\*\*')
	$Items | foreach-Object {
		if ($_.LastUsedTimeStart) {
			$Object = $TemplateObject.PSObject.Copy()
			$Object.Type = 'StoreApp'
			$Object.Name = $($_.PSChildName -split '_')[0]
			$Object.Device = $_.PSParentPath | Split-Path -Leaf
			$Object.LastUsedStartTime = [datetime]::FromFileTime($_.LastUsedTimeStart)
			$Object.LastUsedStopTime = 0
			if ($_.LastUsedTimeStop -ne 0) {
				$Object.LastUsedStopTime = [datetime]::FromFileTime($_.LastUsedTimeStop)
			}
			$Object
		}
	}
}

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
	.EXAMPLE
		Get-DevicePDO -FriendlyName "Some Camera Name"
	#>
	[CmdLetBinding(DefaultParameterSetName='DeviceClass')]
	param(
		[Parameter(Mandatory=$False,ParameterSetName="DeviceClass")]
		[ValidateSet('Camera','AudioEndpoint','Media','Image')]
		[string[]]$DeviceClass=@('Camera','Media','Image'),
		[Parameter(Mandatory=$True,ParameterSetName="FriendlyName")]
		[string]$FriendlyName
	)
	
	if ($PSCmdlet.ParameterSetName -eq 'DeviceClass') {
		foreach ($Class in $DeviceClass) {
			[array]$DeviceClassWMIStrings += "DeviceClass='$Class'"
		}
		[string]$Query = "select * from Win32_PnPSignedDriver where $($DeviceClassWMIStrings -join ' or ')"
	}
	else {
		[string]$Query = "select * from Win32_PnPSignedDriver where friendlyname='$FriendlyName'"
	}

	Write-Debug "WMI Query '$Query'"
	return Get-CimInstance -Query $Query | Select-Object FriendlyName,DeviceClass,PDO
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
	.PARAMETER ProcessID
		Provide the ID of the process if it is known which process has the device in use
	.PARAMETER PDO
		The Physical Device Object name of the device to search the handle for
	.PARAMETER AcceptEula
		Force AcceptEula on Handle.exe
	.EXAMPLE
		Get-DeviceInUseByProcess -Handle .\handle64.exe -PDO \device\00000059
	#>
	[CmdLetBinding(DefaultParameterSetName='__AllParameterSets')]
	param(
		[Parameter(Mandatory=$True)]
		[string]$Handle,
		[Parameter(Mandatory=$False,ParameterSetName='Name')]
		[string[]]$ProcessName,
		[Parameter(Mandatory=$False,ParameterSetName='PID')]
		[int[]]$ProcessID,
		[Parameter(Mandatory=$True)]
		[string]$PDO,
		[Parameter(Mandatory=$False)]
		[switch]$AcceptEula
	)
 
	$AcceptEulaCommandLine = $null
	if ($AcceptEula) {
		$AcceptEulaCommandLine = "/accepteula"
	}

	$Objects = @()
	$HandleInputs = @()
	#Build parameters for handle.exe
	if ($ProcessName) {
		foreach ($ProcName in $ProcessName) {
			$HandleInputs += $ProcName
		}
	}
	elseif ($ProcessID) {
		foreach ($ProcID in $ProcessID) {
			$HandleInputs += $ProcID
		}
	}
	else {
		$HandleInputs += 'all'
	}
	$Regex = "(^\S+)(\W+)(pid:)(\W+)(\d+)(\W+).*"
	foreach ($HandleInput in $HandleInputs) {
		if ($HandleInput -eq 'all' -and !$ProcessID -and !$ProcessName) {
			[array]$HandleData = . $Handle -a -nobanner $AcceptEulaCommandLine
		}
		else {
			Write-Verbose "Getting handles for Process '$HandleInput'"
			[array]$HandleData = . $Handle -p $HandleInput -a -nobanner $AcceptEulaCommandLine
			if ($ProcessID) {
				$Process = Get-Process -Id $HandleInput -erroraction Continue
			}
		}
		for ($i=0;$i-lt$HandleData.Count;$i++) {
			if ($HandleData[$i] -like '----------*') {

			}
			elseif (![string]::IsNullOrWhiteSpace($HandleData[$i])) {
				if ($HandleData[$i] -match $Regex -or $ProcessID) {
					$Object = '' | Select-Object ProcessName,ProcessId,PDO
					if($ProcessID) {
						$Object.ProcessName = $Process.Name
						$Object.ProcessId = $HandleInput
					}
					else {
						$Object.ProcessName = $Matches[1]
						$Object.ProcessId = $Matches[5]
					}
					Write-Debug "Found process '$($Object.ProcessName)' with PID '$($Object.ProcessId)'"
					for ($j=$i+1;$j-lt$HandleData.Count;$j++) {
						if ($HandleData[$j] -match $Regex -and !$ProcessID) {
							Write-Debug "Hit next process with name '$($Matches[1])'"
							$i=$j-1
							break
						}
						elseif ($HandleData[$j] -like "*$PDO") {
							Write-Verbose "Found Physical Device Object name '$PDO' within process '$($Object.ProcessName)'"
							$Object.PDO = $HandleData[$j].Trim()
						}
					}
					if ($Object.PDO) {
						$Objects += $Object
						if($ProcessID) {
							break
						}
					}
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
	.PARAMETER ProcessID
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
	.PARAMETER AcceptEula
		See Get-DeviceInUseByProcess
	.NOTES
		Install-Module -Name Microsoft.PowerShell.SecretsManagement -RequiredVersion 0.2.0-alpha1 -AllowPrerelease
	.EXAMPLE
		Set-HAEntityStateByDeviceInUse -ProcessName svchost -Handle .\handle64.exe -PDO "\Device\00000059" -Uri "http://hassio.local:8123/" -Entity 'input_boolean.in_a_call' -FoundStateValue 'On' -NotFoundStateValue 'Off' -SecretName HAToken -Loop 10000 -verbose
	.EXAMPLE
		Set-HAEntityStateByDeviceInUse -ProcessName svchost,HdxTeams -Handle .\handle64.exe -PDO "\Device\00000059" -Uri "http://hassio.local:8123/" -Entity 'input_boolean.in_a_call' -FoundStateValue 'On' -NotFoundStateValue 'Off' -SecretName HAToken -Loop 10000 -verbose
	#>
	[CmdLetBinding(DefaultParameterSetName='__AllParameterSets')]
	param(
		[Parameter(Mandatory=$True)]
		[string]$Handle,
		[Parameter(Mandatory=$False,ParameterSetName='Name')]
		[string[]]$ProcessName,
		[Parameter(Mandatory=$False,ParameterSetName='PID')]
		[int]$ProcessID,
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
		[switch]$AcceptEula,
		[Parameter(Mandatory=$False)]
		[switch]$Force,
		[Parameter(Mandatory=$False)]
		[int]$Loop
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
	if ($ProcessId) {
		$DeviceInUseByProcessParams['ProcessID'] = $ProcessId
	}
	if ($AcceptEula) {
		$DeviceInUseByProcessParams['AcceptEula'] = $True
	}

	while ($True) {
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
			write-Verbose $($Response | out-string)
		}
		if (!$Loop) {
			break
		}
		start-sleep -milliseconds $Loop
	}
}

Function Set-HAEntityStateByConsentStore {
	<#
	.SYNOPSIS
		Change the state of an entity in HA based on device is in use
	.DESCRIPTION
		Set the state of an entity in HA based on if a device is in use. This is done based on the value of HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\[Device]\NonPackaged.
	.PARAMETER Device
		Provide a device that is registered at HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore. By default de webcam and microphone are used
	.PARAMETER Executable
		One or more executables that will be monitored
	.PARAMETER ExcludeExecutable
		Enable to exclude executables. If -Executable is provided, those will be excluded.
	.PARAMETER StoreApp
		One or more store apps can be added to include in the monitoring
	.PARAMETER ExcludeStoreApp
		Enable if you want to exclude the store apps. If -StoreApps has got values, those will be excluded.
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
	.EXAMPLE
		Set-HAEntityStateByConsentStore -executable speechrecognition.exe -excludeexecutable -Uri "http://hassio.local:8123/" -Entity 'input_boolean.in_a_call' -FoundStateValue 'On' -NotFoundStateValue 'Off' -SecretName HAToken -verbose
	.EXAMPLE
		Set-HAEntityStateByConsentStore -Uri "http://hassio.local:8123/" -Entity 'input_boolean.in_a_call' -FoundStateValue 'On' -NotFoundStateValue 'Off' -SecretName HAToken -verbose
	#>
	[CmdLetBinding(DefaultParameterSetName='__AllParameterSets')]
	Param(
		[Parameter(Mandatory=$False)]
		[string[]]$Device=@('microphone','webcam'),
		[Parameter(Mandatory=$False)]
		[string[]]$Executable,
		[Parameter(Mandatory=$False)]
		[switch]$ExcludeExecutable,
		[Parameter(Mandatory=$False)]
		[string[]]$StoreApp,
		[Parameter(Mandatory=$False)]
		[switch]$ExcludeStoreApp,
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

	if ($ExcludeExecutable -and $ExcludeStoreApp -and !$Executable -and !$StoreApp) {
		Write-Error "-ExcludeExecutable and -ExcludeStoreApp are defined without -Executable and/or -StoreApp. This will not ennumerate any programs" -ErrorAction Stop
	}

	#Get the secret
	$SecretParams = @{}
	if ($VaultName) {
		$SecretParams['Vault'] = $VaultName
	}
	[string]$LongLivedAccessToken = get-secret -Name $SecretName -AsPlainText @SecretParams -erroraction stop
	Write-Debug "Got token '$($LongLivedAccessToken)'"

	$endpoint = "$($uri.AbsoluteUri)api/states/$Entity"
	write-debug "Endpoint will be '$endpoint'"

	$Drive = $(get-psdrive).where({$_.Name -eq 'hku'})
	if (!$Drive) {
		Write-Verbose "'HKU' is not mapped, mapping it now"
		New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | out-null
	}
	
	$objuser = New-Object System.Security.Principal.NTAccount($env:username)
	$SID = $objuser.Translate([System.Security.Principal.SecurityIdentifier]).Value
	$RootPath = "$SID\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore"
	$FullPath = join-path 'HKU:' $RootPath
	
	if (!$(test-path $FullPath)) {
		Throw "'$FullPath' does not exist"
	}

	#Register the wait for change last, so on launch the state can be set correct
	$Query = "Select * from RegistryTreeChangeEvent WHERE Hive='HKEY_USERS' AND RootPath='$($RootPath -replace '\\','\\')'"
	$SourceIdentifier = 'ConsentStore'

	$First = $True
	while ($True) {
		$StateToSend = $NotFoundStateValue
		$TriggerApp = $null

		$Path = Join-Path 'HKU:' $SID
		
		for ($j = 0; $j -lt $Device.count -and $null -eq $TriggerApp; $j++) {
			$DevicePath = join-path $FullPath $Device[$j]
			$NonPackaged = join-path $DevicePath 'NonPackaged'
			$StoreApps = join-path $DevicePath '*'
			$Items = @()
			if ($(test-path $DevicePath)) {
				if (!$(!$Executable -and $ExcludeExecutable)) {
					Write-Debug "Ennumerating NonPackages executables registry '$NonPackaged'"
					if ($(test-path $NonPackaged)) {
						$NonPackegedValues = join-path $NonPackaged '*'
						write-Debug  "Getting info from '$($NonPackegedValues)'"
						$Items += $(Get-ItemProperty $NonPackegedValues -erroraction silentlycontinue | Where-Object {$_.LastUsedTimeStart -and $_.LastUsedTimeStop -eq 0})
					}
				}
				else {
					Write-Debug "Nonpackaged is skipped with -ExcludeExectuables"
				}

				if (!$(!$StoreApp -and $ExcludeStoreApp)) {
					Write-Debug "Ennumerating store app registry '$StoreApps'"
					$Items += $(Get-ItemProperty $StoreApps | Where-Object {$_.LastUsedTimeStart -and $_.LastUsedTimeStop -eq 0})
				}
				else {
					Write-Debug "Store apps are skipped with -ExcludeStoreApps"
				}
			}

			if ($Items) {
				for ($k = 0; $k -lt $Items.count -and $null -eq $TriggerApp; $k++) {
					write-debug "Accessing '$($Items[$k].PSPath)'"
					#Get the executable name from the registry key
					if ($Items[$k].PSPath -like '*\NonPackaged\*') {
						Write-Debug "NonPackaged '$($Items[$k].PSChildName)' is using the '$($Device[$j])' since '$([datetime]::FromFileTime($Items[$k].LastUsedTimeStart))'"
						$DeviceExecutable = $($Items[$k].PSChildName -split '#')[-1]
						if ($ExcludeExecutable -and $DeviceExecutable -in $Executable) {
							Write-Debug "Found '$DeviceExecutable' using the device '$($Device[$j])', but it is excluded"
						}
						elseif (!$Executable -or (!$ExcludeExecutable -and $DeviceExecutable -in $Executable) -or ($ExcludeExecutable -and $DeviceExecutable -notin $Executable)) {
							Write-Debug "Found '$DeviceExecutable' using the device '$($Device[$j])'"
							$StateToSend = $FoundStateValue
							$TriggerApp = $DeviceExecutable
						}
					}
					else {
						Write-Debug "StoreApp '$($Items[$k].PSChildName)' is using the '$($Device[$j])' since '$([datetime]::FromFileTime($Items[$k].LastUsedTimeStart))'"
						$StoreAppName = $($Items[$k].PSChildName -split '_')[0]
						if ($ExcludeStoreApp -and $StoreAppName -in $StoreApp) {
							Write-Debug "Found '$StoreAppName' using the device '$($Device[$j])', but it is excluded"
						}
						elseif (!$StoreApp -or (!$ExcludeStoreApp -and $StoreAppName -in $StoreApp) -or ($ExcludeStoreApp -and $StoreAppName -notin $StoreApp)) {
							Write-Debug "Found '$StoreAppName' using the device '$($Device[$j])'"
							$StateToSend = $FoundStateValue
							$TriggerApp = $StoreAppName
						}
					}
				}
			}
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
			Write-Verbose "The state should be changed to '$StateToSend', triggered by '$TriggerApp'"
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
			write-Verbose $($Response | out-string)
		}

		Register-WMIEvent -Query $Query -SourceIdentifier $SourceIdentifier -ErrorAction SilentlyContinue
		
		Wait-Event -SourceIdentifier $SourceIdentifier -Timeout 30 | out-null
		Remove-Event -SourceIdentifier $SourceIdentifier -ErrorAction SilentlyContinue | out-null

	}
	Unregister-Event -SourceIdentifier $SourceIdentifier | out-null
}
