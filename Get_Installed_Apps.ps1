function Get-InstalledApplications {
    [cmdletbinding(DefaultParameterSetName = 'GlobalAndAllUsers')]

 

    Param (
        [Parameter(ParameterSetName="Global")]
        [switch]$Global,
        [Parameter(ParameterSetName="GlobalAndCurrentUser")]
        [switch]$GlobalAndCurrentUser,
        [Parameter(ParameterSetName="GlobalAndAllUsers")]
        [switch]$GlobalAndAllUsers,
        [Parameter(ParameterSetName="CurrentUser")]
        [switch]$CurrentUser,
        [Parameter(ParameterSetName="AllUsers")]
        [switch]$AllUsers
    )

 

    # Obtener el nombre del equipo
    $PC = $env:COMPUTERNAME
    #Write-Host "Nombre del equipo: $ComputerName"

 

    # Excplicitly set default param to True if used to allow conditionals to work
    if ($PSCmdlet.ParameterSetName -eq "GlobalAndAllUsers") {
        $GlobalAndAllUsers = $true
    }

 

    # Check if running with Administrative privileges if required
    if ($GlobalAndAllUsers -or $AllUsers) {
        $RunningAsAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($RunningAsAdmin -eq $false) {
            Write-Error "Finding all user applications requires administrative privileges"
            break
        }
    }

 

    # Empty array to store applications
    $Apps = @()
    $32BitPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $64BitPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"

    #Retrieve Appx Packages
    $Apps += Get-AppxPackage -Allusers  | Select @{N='displayName';E={$_.Name}},@{N='DisplayVersion';E={$_.Version}},@{N='Publisher';E={$_.Publisher.split(',')[0].remove(0,3)}}

    # Retrieve globally installed applications
    if ($Global -or $GlobalAndAllUsers -or $GlobalAndCurrentUser) {
        #Write-Host "Processing global hive"
        $Apps += Get-ItemProperty "HKLM:\$32BitPath"
        $Apps += Get-ItemProperty "HKLM:\$64BitPath"
    }

 

    if ($CurrentUser -or $GlobalAndCurrentUser) {
        #Write-Host "Processing current user hive"
        $Apps += Get-ItemProperty "Registry::\HKEY_CURRENT_USER\$32BitPath"
        $Apps += Get-ItemProperty "Registry::\HKEY_CURRENT_USER\$64BitPath"
    }

 

    if ($AllUsers -or $GlobalAndAllUsers) {
        #Write-Host "Collecting hive data for all users"
        $AllProfiles = Get-CimInstance Win32_UserProfile | Select LocalPath, SID, Loaded, Special | Where-Object {$_.SID -like "S-1-5-21-*" -or $_.SID -like "S-1-12-*"}
        $MountedProfiles = $AllProfiles | Where-Object {$_.Loaded -eq $true}
        $UnmountedProfiles = $AllProfiles | Where-Object {$_.Loaded -eq $false}

 

        #Write-Host "Processing mounted hives"
        $MountedProfiles | ForEach-Object {
            $Apps += Get-ItemProperty -Path "Registry::\HKEY_USERS\$($_.SID)\$32BitPath"
            $Apps += Get-ItemProperty -Path "Registry::\HKEY_USERS\$($_.SID)\$64BitPath"
        }

 

        #Write-Host "Processing unmounted hives"
        $UnmountedProfiles | ForEach-Object {

 

            $Hive = "$($_.LocalPath)\NTUSER.DAT"
            #Write-Host " -> Mounting hive at $Hive"

 

            if (Test-Path $Hive) {
                REG LOAD HKU\temp $Hive > $null

 

                $Apps += Get-ItemProperty -Path "Registry::\HKEY_USERS\temp\$32BitPath"
                $Apps += Get-ItemProperty -Path "Registry::\HKEY_USERS\temp\$64BitPath"

 

                # Run manual GC to allow hive to be unmounted
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()

 

                REG UNLOAD HKU\temp > $null
            } else {
                Write-Warning "Unable to access registry hive at $Hive"

 

            }
            
        }
    }

 

    # Definir los nuevos nombres de las propiedades
    $Properties = @{
        "DisplayName" = "Name"
        "DisplayVersion" = "Ver"
        "InstallDate" = "Date"
        "Publisher" = "Pub"
#        "InstallLocation" = "Location"
#        "QuietUninstallString",
#        "UninstallString",
#        "EstimatedSize",
#        "Language" = "Lang"
#        "PSPath",
#        "PSParentPath",
        "PSChildName" = "ChName"
#        "PSProvider"
    }



    # Crear un objeto personalizado con el nombre del equipo y las aplicaciones instaladas
    $CustomObjects = $Apps | ForEach-Object {
        $CustomObject = New-Object PSObject
        $CustomObject | Add-Member -NotePropertyName "PC" -NotePropertyValue $PC
        foreach ($property in $Properties.GetEnumerator()) {
            $propertyName = $property.Name
            if ($_.PSObject.Properties.Name -contains $propertyName) {
                $CustomObject | Add-Member -NotePropertyName $property.Value -NotePropertyValue $_.$propertyName
            } else {
                $CustomObject | Add-Member -NotePropertyName $property.Value -NotePropertyValue ""
            }
        }
        $CustomObject 
    }
     
    # Mostrar la salida en formato CSV con las propiedades definidas
    $CustomObjects = $CustomObjects | Group-Object -Property Name, Ver | ForEach-Object {
    # Si hay más de un objeto en el grupo, seleccionar el primero
    if ($_.Count -gt 1) {
        $_.Group[0]
    }
    else {
        $_.Group
    }
}

# Mostrar la salida en formato CSV con las propiedades definidas
$CustomObjects | ConvertTo-Csv #| ConvertTo-Json -Compress
}

Get-InstalledApplications
