# HomeAssistantDeviceState
This module allows you to set an Home Assistant entity based on the usage of a device within in Windows. It is mainly developed for detecting camera usage within teams, but microphones can also be detected. Handle.exe from sysinternals is used to detect a handle of a process to the specified device. The module has been tested on Powershell 5 and 7.

## Getting started
Download handle.exe from https://docs.microsoft.com/en-us/sysinternals/downloads/handle. After downloading you should execute one of the executables to accept the EULA.

Launch Powershell and execute the following command, make sure PowershellGet is up-to-date
```Powershell
Install-Module -Name Microsoft.PowerShell.SecretsManagement -RequiredVersion 0.2.0-alpha1 -AllowPrerelease
```

Go to Home assistant and generate an Long Lived Access token, go back to Powershell and type the following to store the token securely
```Powershell
Add-Secret -name HAToken -Secret "Your Long Lived Access token"
```

Download the module and enter this command in Powershell
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

Create a ps1 file with the following contents. Make sure that all parameters are correct. The example below changes an input boolean from on to off and vice versa.

```Powershell
start-transcript "$PSScriptRoot\HomeAssistantDeviceState.log"
$VerbosePreference = 'Continue'

Import-Module $PSScriptRoot\HomeAssistantDeviceState.psd1

while ($true) {
    Write-Verbose "Starting @ '$(get-date)'"
    Set-HAEntityStateByDeviceInUse -ProcessName svchost -Handle .\handle64.exe -PDO "\Device\00000060" -Uri "http://hassio.local:8123/" -Entity 'input_boolean.camera' -FoundStateValue 'On' -NotFoundStateValue 'Off' -SecretName HAToken -verbose
    start-sleep -Seconds 10
}

stop-transcript
```

This ps1 looks for \Device\00000060 in the process svchost. The ProcessName can be left out, but it is wise to limit this to a specific ProcessName. Without any ProcessName specified, Handle.exe uses a lot of process power. Multiple processes can be added seperated by a comma. If you'll need to find another process, you can use the command below while your device is in use.
```Powershell
 Get-DeviceInUseByProcess -Handle .\handle64.exe -PDO \device\0000600 -Verbose
 ```

The ps1 can be used to be started through the task scheduler. Point the action to powershell/pwsh with the arguments below
```
-nop -exec bypass -command "& 'Path\To\PowershellScript.ps1'"
```
A trigger that works is "At logon" with "Repeat the task every 1 minute". The script keeps on running in the backgrond, but if it crashes for some reason, it will be started again.

If you do not point to the exact location of handle.exe, and it is in the same directory as your ps1, make sure the "Start in" configuration in the scheduled task points to the directory of you ps1.
