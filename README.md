# HomeAssistantDeviceState
This module allows you to set an Home Assistant entity based on the usage of a device within in Windows. It is mainly developed for detecting camera usage within teams, but microphones can also be detected. The module has been tested on Powershell 5 and 7.

## Getting started

Launch Powershell and execute the following command, make sure PowershellGet is up-to-date
```Powershell
Install-Module -Name Microsoft.PowerShell.SecretsManagement -RequiredVersion 0.2.0-alpha1 -AllowPrerelease
```

Go to Home assistant and generate an Long Lived Access token, go back to Powershell and type the following to store the token securely
```Powershell
Add-Secret -name HAToken -Secret "Your Long Lived Access token"
```

You will need an entity in Home Assistant, below there is an input_boolean entity that you can add. It is possible to have another type of entity, as long as the value can be set with -FoundValueState and -NotFoundValueState at the command Set-HAEntityStateByDeviceInUse. The command Set-HAEntityStateByDeviceInUse has only been tested with an input boolean.
```yaml
input_boolean:
  in_a_call:
    name: In a call
    initial: off
    icon: mid:webcam
```

There are two methods 2 use, the ContentStore registry information and Handle.exe. For me, the ContentStore method suffices in retreiving the Camera and Microphone status. The Handle.exe method, i never got it to work with the microphone. The best method I think is the ContentStore method

# ContentStore registry

Create a ps1 file with the following contents beside the downloaded module. If you want to monitor explicit for one or more executables add it at -Executable, add -Exclude to exclude the -Executable from being monitord. The example below monitors for every executable.

```Powershell
start-transcript "$PSScriptRoot\HomeAssistantDeviceState.log"
$VerbosePreference = 'Continue'

Import-Module $PSScriptRoot\HomeAssistantDeviceState.psd1

Set-HAEntityStateByConsentStore -Uri "http://hassio.local:8123/" -Entity 'input_boolean.in_a_call' -FoundStateValue 'On' -NotFoundStateValue 'Off' -SecretName HAToken -verbose

stop-transcript
```

# Shortcut
The ConentStore method can best be started through a shortcut in the startup folder, use c:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\ if you want to start it for all users. Create a new shortcut and point it to
```
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -windowstyle hidden -nop -exec bypass -command "& 'Path\To\PowershellScript.ps1'"
```
Only Powershell 5 can be used.

# Handle.exe
Download handle.exe from https://docs.microsoft.com/en-us/sysinternals/downloads/handle. After downloading you should execute one of the executables to accept the EULA. The CmdLets that require handle.exe as parameter, also have -accepteula as parameter. Every user that runs handle.exe should also accept the sysinternals eula.

Enter this command in Powershell
```Powershell
Import-Module Path\To\HomeAssistantDeviceState.psd1
```

Now you will need to know what the Physical Device Object name is. This can be done by just executing Get-DevicePDO
```Powershell
Get-DevicePDO

FriendlyName                DeviceClass PDO
------------                ----------- ---
Imaging device              IMAGE       \Device\0000015c
                            MEDIA       \Device\00000053
Some Camere Name Camera     CAMERA      \Device\00000060

#If the above does not list the device that will be monitord, find the friendly name of your camera in Device Manager and try this command
Get-DevicePDO -friendlyname 'Camera friendly name'
```

Create a ps1 file with the following contents beside the downloaded module. Make sure that all parameters are correct. The example below changes an input boolean from on to off and vice versa.

```Powershell
start-transcript "$PSScriptRoot\HomeAssistantDeviceState.log"
$VerbosePreference = 'Continue'

Import-Module $PSScriptRoot\HomeAssistantDeviceState.psd1

#Sometimes the PDO changes after a reboot, use this to always have the right PDO
$DevicePDO = $(Get-DevicePDO).where({$_.FriendlyName -eq 'Some Camera Name' -and $_.DeviceClass -eq 'Camera'})

Set-HAEntityStateByDeviceInUse -ProcessName svchost -Handle .\handle64.exe -PDO $DevicePDO.PDO -Uri "http://hassio.local:8123/" -Entity 'input_boolean.camera' -FoundStateValue 'On' -NotFoundStateValue 'Off' -SecretName HAToken -Loop 10000 -verbose

stop-transcript
```

This ps1 looks for \Device\00000060 in the process svchost. The ProcessName can be left out, but it is wise to limit this to a specific ProcessName. Without any ProcessName specified, Handle.exe uses a lot of process power. Multiple processes can be added seperated by a comma. If you'll need to find another process, you can use the command below while your device is in use.
```Powershell
 Get-DeviceInUseByProcess -Handle .\handle64.exe -PDO \device\0000600 -Verbose
 ```

# Handle scheduled task
The ps1 created for the Handle method can be used to start through the task scheduler. Point the action to powershell/pwsh with the arguments below and run in as the same user that will use the device. You can also run it as SYSTEM, but then you will have to add the secret as SYSTEM. This can be done with psexec, or you can run the .ps1 one time as SYSTEM with Add-Secret.
```
-nop -exec bypass -command "& 'Path\To\PowershellScript.ps1'"
```
A trigger that works is "At logon" with "Repeat the task every 1 minute". The script keeps on running in the backgrond, but if it crashes for some reason, it will be started again.

When using the handle.exe method: if you do not point to the exact location of handle.exe, and it is in the same directory as your ps1, make sure the "Start in" configuration in the scheduled task points to the directory of your ps1.
