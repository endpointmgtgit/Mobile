<#
  .SYNOPSIS
  Allows you to initiate a device wipe/retire and removal from Microsoft Intune based on a csv file of users.

  .DESCRIPTION
  The Invoke-DeviceWipe.ps1 allows for bulk removal of users devices from Intune based on username associated with the device.

  .PARAMETER csvFile
  Path to the csv file containing a list of user user principal names, or a csv as an output from this script.

  .PARAMETER operatingSystem
  Selection of operating systems from the currently supporting Microsoft Intune device management objects.
  Valid options: Android

  .PARAMETER deviceOwner
  Select whether you are capturing company or personal owned managed devices.
  Valid options: company, personal

  .PARAMETER wipeDevice
  Specifies whether the script will initiate a device wipe on the selected devices.
  Valid options: Yes, No.

  If No and with a csv of only user principal names, will export a csv of the users devices. Recommended Step 1
  If Yes and with a csv of only user principal names, will wipe devices on the fly. Not recommended

  If Yes and with a csv of the exported devices, will wipe the devices from the csv file. Recommended Step 2
  If No and with a csv of the exported devices, will end the script.

  .INPUTS
  None. You can't pipe objects to Invoke-DeviceWipe.ps1

  .OUTPUTS
  None. Invoke-DeviceWipe.ps1 doesn't generate any output.

  .EXAMPLE
  PS> ./Invoke-DeviceWipe.ps1 -csvFile C:\source\users.csv -operatingSystem Android -deviceOwner company -wipeDevice No

  .EXAMPLE
  PS> ./Invoke-DeviceWipe.ps1 -csvFile 'C:\source\users devices to delete.csv' -operatingSystem Android -deviceOwner company -wipeDevice Yes
#>

[CmdletBinding()]

param(

    [Parameter(Mandatory = $true)]
    [String]$csvFile,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Android')]
    [String]$operatingSystem,

    [Parameter(Mandatory = $true)]
    [ValidateSet('company', 'personal')]
    [String]
    $deviceOwner,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Yes', 'No')]
    [String]$wipeDevice = 'No'

)

#region parameters
$rndWait = Get-Random -Minimum 3 -Maximum 10
$Scopes = 'Device.ReadWrite.All,DeviceManagementManagedDevices.ReadWrite.All,DeviceManagementConfiguration.ReadWrite.All'
if (!(Test-Path $csvFile -PathType Leaf)) {
    Write-Host "Unable to find the specified csv file $csvFile please re-run the script." -ForegroundColor Red
    Break
}
else {
    $users = Import-Csv -Path $csvFile
    if ($users.count -eq 0 -or $null -eq $users) {
        Write-Host "Unable to find users specified csv file $csvFile please re-run the script." -ForegroundColor Red
        Break
    }
    else {
        if ($null -eq $users.id) {
            Write-Host "The CSV file $csvFile only contains users, the script will capture the device objects from Microsoft Intune." -ForegroundColor Yellow
            Write-Host
            $runMode = 'OnTheFly'

        }
        else {
            Write-Host "The CSV file $csvFile contains users and associated device objects, the script will initate a device wipe from Microsoft Intune" -ForegroundColor Yellow
            Write-Host
            $runMode = 'FromCSV'
        }
    }
}
#endregion parameters

#region authentication
$moduleName = 'Microsoft.Graph'
$Module = Get-InstalledModule -Name $moduleName
if ($Module.count -eq 0) {
    Write-Host "$moduleName module is not available" -ForegroundColor yellow
    $Confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No
    if ($Confirm -match '[yY]') {
        Install-Module -Name $moduleName -AllowClobber -Scope CurrentUser -Force
    }
    else {
        Write-Host "$moduleName module is required. Please install module using 'Install-Module $moduleName -Scope CurrentUser -Force' cmdlet." -ForegroundColor Yellow
        break
    }
}
else {
    If ($IsMacOS) {
        Connect-MgGraph -Scopes $Scopes -UseDeviceAuthentication -ContextScope Process
    }
    If ($IsWindows) {
        Connect-MgGraph -Scopes $Scopes -UseDeviceCode
    }
}
#endregion authentication

#region script
if ($runMode -eq 'OnTheFly') {
    Write-Host "Getting $operatingSystem devices from Microsoft Intune." -ForegroundColor Cyan
    Write-Host
    $managedDevices = Get-MgDeviceManagementManagedDevice | Where-Object { ($_.OperatingSystem -eq $operatingSystem) -and ($_.ManagedDeviceOwnerType -eq $deviceOwner) }
    if ($null -eq $managedDevices -or $managedDevices.count -eq 0) {
        Write-Host "No $operatingSystem devices found devices from Microsoft Intune." -ForegroundColor Red
        Break
    }
    Write-Host "Found $($managedDevices.count) $operatingSystem devices from Microsoft Intune." -ForegroundColor Green
    Write-Host

    $manageddevicesToWipe = @()
    foreach ($user in $users) {
        Write-Host "Getting $operatingSystem devices from Microsoft Intune for $($user.UserPrincipalName)" -ForegroundColor Cyan
        Write-Host
        $userDevices = $managedDevices | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName }
        if ($null -eq $userDevices -or $userDevices.count -eq 0 ) {
            Write-Host "No $operatingSystem devices from Microsoft Intune for $($user.UserPrincipalName)" -ForegroundColor Yellow
            Write-Host
        }
        else {
            foreach ($userDevice in $userDevices) {
                $manageddevicesToWipe += [pscustomobject]@{User = $userDevice.UserDisplayName; UserPrincipalName = $user.UserPrincipalName; Name = $userDevice.DeviceName; Id = $userDevice.Id; Manufacturer = $userDevice.Manufacturer; Model = $userDevice.Model }
            }
            Write-Host "Found $($userDevices.count) $operatingSystem devices from Microsoft Intune for $($user.UserPrincipalName)" -ForegroundColor Green
            Write-Host
        }
    }

    if ($wipeDevice -eq 'Yes') {
        Write-Host 'Please review the devices that are to be removed from Microsoft Intune' -ForegroundColor Yellow
        $manageddevicesToWipe.Name
        Write-Warning 'Please confirm you are happy to continue to wipe the list of devices from Microsoft Intune.' -WarningAction Inquire
        foreach ($manageddeviceToWipe in $manageddevicesToWipe) {
            Write-Host "Removing device $($manageddeviceToWipe.Name), $($manageddeviceToWipe.Manufacturer) $($manageddeviceToWipe.Model) from Microsoft Intune for $($manageddeviceToWipe.User)" -ForegroundColor Yellow
            Write-Host
            Start-Sleep -Seconds $rndWait
            try {
                Remove-MgDeviceManagementManagedDevice -ManagedDeviceId $manageddeviceToWipe.Id
                Write-Host 'Removed device from Microsoft Intune.' -ForegroundColor Green
            }
            catch {
                Write-Host 'Unable to remove device from Microsoft Intune.' -ForegroundColor Red
            }
        }
    }
    else {
        $csvExport = (Split-Path -Path $csvFile) + '/' + (Get-Date -Format yyyyMMddhhmmss) + '-devicesToWipe.csv'
        $manageddevicesToWipe | Export-Csv -Path $csvExport
        Write-Host "A list of devices has been created in $csvExport, please re-run the script- and use this csv file as the source." -ForegroundColor Green
        Break
    }
}
else {
    if ($wipeDevice -eq 'Yes') {
        Write-Host
        Write-Host 'Please review the devices that are to be removed from Microsoft Intune' -ForegroundColor Yellow
        $users.Name
        Write-Warning 'Please confirm you are happy to continue to wipe the list of devices from Microsoft Intune.' -WarningAction Inquire
        foreach ($manageddeviceToWipe in $users) {
            Write-Host "Removing $operatingSystem device $($manageddeviceToWipe.Name), $($manageddeviceToWipe.Manufacturer) $($manageddeviceToWipe.Model) from Microsoft Intune for $($manageddeviceToWipe.User)" -ForegroundColor Yellow
            Write-Host
            Start-Sleep -Seconds $rndWait
            try {
                Remove-MgDeviceManagementManagedDevice -ManagedDeviceId $manageddeviceToWipe.Id
                Write-Host 'Removed device from Microsoft Intune.' -ForegroundColor Green
            }
            catch {
                Write-Host 'Unable to remove device from Microsoft Intune.' -ForegroundColor Red
            }
            Write-Host
        }
    }
    else {
        Write-Host 'Re-run the script when you are ready to remove devices from Microsoft Intune.' -ForegroundColor Red
        Break
    }
}

#endregion script