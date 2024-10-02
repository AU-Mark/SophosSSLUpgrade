Try {
    $CurrentUser = Get-WmiObject Win32_Process -Filter "Name='explorer.exe'" | ForEach-Object { $_.GetOwner() } | Select-Object -Unique -Expand User

    # Retrieve the current VPN profile
    $OVPN = @(Get-ChildItem -Path "C:\Program Files (x86)\Sophos\Sophos SSL VPN Client\config" -Filter "$CurrentUser*.ovpn" -File)[0]

    If (!$OVPN) {
        Write-Host "No OVPN file could be located for $CurrentUser"
        $OVPNFound = @(Get-ChildItem -Path "C:\Program Files (x86)\Sophos\Sophos SSL VPN Client\config" -Filter "*.ovpn" -File)[0]
        If ($OVPNFound) {
            $Response = Read-Host "A OVPN File was found with filename $($OVPNFound.Name). Would you like to use this OVPN file for $CurrentUser? (Y/N)"
            If ($Response.ToUpper() -eq "Y") {
                $OVPN = @(Get-ChildItem -Path "C:\Program Files (x86)\Sophos\Sophos SSL VPN Client\config" -Filter "*.ovpn" -File)[0]
            }
        }
    }
    
    If ($OVPN) {
        # Set the temp VPN profile location we will be using
        $TempOVPN = "C:\Temp\$($OVPN.Name)"
    }

    #Variable for if OVPN file was found
    $OVPNFound = $False

    # Make changes to the OVPN profile to match the options as if downloaded a new OVPN profile for the user
    If ($OVPN) {
        $OVPNFound = $True
        $VPNProfile = Get-Content $OVPN.FullName
        $VPNProfile = $VPNProfile.Replace('ip-win32 dynamic','')
        $VPNProfile = $VPNProfile.Replace('comp-lzo no','comp-lzo yes')
        # These seem to be default values but could possibly be different based on the VPN user...
        $VPNProfile = $VPNProfile += (";can_save no`n;otp no`n;run_logon_script no`n;auto_connect")
        $VPNProfile > $TempOVPN
    } 

    # Uninstall the old SSL VPN client silently
    If (Test-Path "C:\Program Files (x86)\Sophos\Sophos SSL VPN Client\Uninstall.exe") {
        Write-Host "Uninstalling Sophos SSL VPN Client"
        Start-Process -Wait -FilePath "C:\Program Files (x86)\Sophos\Sophos SSL VPN Client\Uninstall.exe" -ArgumentList "/S /qn"
        Write-Host "Uninstalled Sophos SSL VPN Client"
    }

    # Download and silently install Sophos VPN Connect client
    Write-Host "Downloading the latest Sophos Connect VPN client from Sophos"
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri "https://github.com/AU-Mark/SophosSSLUpgrade/raw/refs/heads/main/SophosConnect_2.3.1_IPsec_and_SSLVPN.msi" -OutFile "C:\Temp\SophosConnect_2.3.1.msi"
    Write-Host "Installing Sophos Connect"
    Start-Process msiexec.exe -Wait -ArgumentList '/i "C:\Temp\SophosConnect_2.3.1.msi" /quiet /qn /norestart'
    Write-Host "Sophos Connect Successfully Installed!"

    If ($OVPNFound) {
        # Install the VPN profile
        If (Test-Path "C:\Program Files (x86)\Sophos\Connect\sccli.exe") {
            Write-Host "Sophos Connect CLI Executable detected, installing VPN profile"

            Try {
            # Install the VPN profile without opening a new window. Otherwise the user will see a flash of a command prompt window on their screen for a brief moment.
            Start-Process -FilePath "C:\Program Files (x86)\Sophos\Connect\sccli.exe" -ArgumentList "add -f $TempOVPN"
            Remove-Item -Path $TempOVPN > $Null

            Write-Host "Installed VPN profile"
            Write-Host "Sophos Connect VPN is all set on this device!"
            } Catch {
                Write-Host "An error occured. See the error line below for more details."
                Write-Error "Err Line: $($_.InvocationInfo.ScriptLineNumber) Err Name: $($_.Exception.GetType().FullName) Err Msg: $($_.Exception.Message)"
            }
        }
    } Else {
        Write-Host "No OVPN file was found. You will need to work with the user to download a new OVPN profile from the firewalls VPN portal and install it to Sophos Connect"
    }
} Catch {
    Write-Host "An error occured. See the error line below for more details."
    Write-Error "Err Line: $($_.InvocationInfo.ScriptLineNumber) Err Name: $($_.Exception.GetType().FullName) Err Msg: $($_.Exception.Message)"
}
