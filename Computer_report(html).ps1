#---------------------------------------------------------------------
# Creating the file
#---------------------------------------------------------------------
$ComputerName = $env:computername

#---------------------------------------------------------------------
# Get Computer info
#---------------------------------------------------------------------

$csinfo = Get-ComputerInfo | Select-Object @{Name="Device Name";Expression={$_.CsCaption}}, @{Name="Installed Ram";Expression={$_.OsTotalVisibleMemorySize/1MB}}, @{Name="ProductID";Expression={$_.WindowsProductId}}

$processor = Get-CimInstance -Class Win32_Processor | Select-Object -Property Name

$csinfo | Add-Member -NotePropertyName "Processor" -NotePropertyValue $processor.Name -Force

$DiskSize = Get-Volume | Select-Object DriveLetter, FileSystem, FileSystemLabel, @{Name="Total Size (GB)";Expression={$_.Size/1GB}}, @{Name="Remaining Size (GB)";Expression={$_.SizeRemaining/1GB}}

#---------------------------------------------------------------------
# Collect operating system information and convert to HTML fragment
#---------------------------------------------------------------------

$osinfo = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction STOP | Select-Object @{Name="Edition";Expression={$_.Caption}},Version,@{Name="OS Build";Expression={$_.BuildNumber}},@{Name="Architecture";Expression={$_.OSArchitecture}}
#--------------------------------------------------------------------
# Html Info
#---------------------------------------------------------------------
$htmlreport = @()
$htmlbody = @()
$username = $env:username
$serial = (Get-WmiObject win32_bios).SerialNumber
$htmlfile = "C:\$serial-$username.html"

#---------------------------------------------------------------------
# Get Administrator account info
#---------------------------------------------------------------------
$Users = Get-CimInstance -Class Win32_UserAccount -Filter "LocalAccount=True" | Select-Object Name, Status, Disabled, PasswordExpires, SID


#---------------------------------------------------------------------
# Get Network info
#---------------------------------------------------------------------

$NetAdapter = Get-NetAdapter | Select-Object Name, Status, DriverVersion, DriverProvider

$VPNstatus = Get-VpnConnection | Select-Object Name, ServerAddress, SplitTunneling, AllUserConnection, L2tpIPsecAuth, ConnectionStatus

# Get network configuration information
$networkConfig = Get-NetIPConfiguration

# Create a list to store network interface data
$networkInterfaces = @()

# Foreach at the network configuration to get the necessary data
foreach ($config in $networkConfig) {
    $interface = [PSCustomObject]@{
    Interface = $config.InterfaceAlias
    IPAddress = $config.IPv4Address.IPAddress
    SubnetMask = $config.IPv4Address.PrefixLength
}
    $networkInterfaces += $interface
}

#---------------------------------------------------------------------
# Get Bitlocker info
#---------------------------------------------------------------------

$Bitlocker = Get-BitLockerVolume -MountPoint 'C:' | Select-Object MountPoint, EncryptionMethod, VolumeStatus, ProtectionStatus, LockStatus, EncryptionPercentage, @{Name="Capacity (GB)";Expression={$_.Capacity/1GB}}

$BitlockerKey = (Get-BitLockerVolume -MountPoint C).KeyProtector | Where-Object { $_.KeyProtectorType -notlike "Tpm" } | Select-Object KeyProtectorId, KeyProtectorType, RecoveryPassword

#---------------------------------------------------------------------
# Create a format to HTML body
#---------------------------------------------------------------------

$htmlhead = @"
<html>
<head>
<style>
    body {
        font-family: Arial, sans-serif;
        font-size: 12px;
        line-height: 1.4;
        color: #333333;
    }

    h1 {
        font-size: 30px;
        font-weight: bold;
        margin-bottom: 10px;
        text-align: center;
    }

    h2 {
        font-size: 26px;
        font-weight: bold;
        color: red;
        margin-bottom: 8px;
    }

    h3 {
        font-size: 22px;
        font-weight: bold;
        margin-bottom: 6px;
    }

    table {
        border-collapse: collapse;
        width: 100%;
    }

    th, td {
        padding: 8px;
        text-align: left;
        border-bottom: 1px solid #dddddd;
    }

    th {
        background-color: #f5f5f5;
    }

    .spacer {
        margin-bottom: 12px;
    }
</style>
<title>Report - $serial</title>
</head>
<body>
"@

$htmltail = @"
</body>
</html>
"@

$htmlbody += @"
</br>
<h1 style="text-align: center; margin: 0 auto;">Laptop $serial</h1>
</br>
</br>
<h2>Computer Info</h2>
<h3>Device Specifications</h3>
"@
$htmlbody += ($csinfo | ConvertTo-Html -Property "Device Name", "Processor", "Installed Ram", "ProductID" -As Table) -replace "<table>", "<table class=`"info-table`">"

$htmlbody += @"
<h3>Disk Partitions</h3>
"@
$htmlbody += ($DiskSize | ConvertTo-Html -As Table) -replace "<table>", "<table class=`"info-table`">"

$htmlbody += @"
<h3>Windows Specifications</h3>
"@
$htmlbody += ($osinfo | ConvertTo-Html -As Table) -replace "<table>", "<table class=`"info-table`">"

$htmlbody += @"
<div class='spacer'></div>

<h2>Local Users</h2>
"@
$htmlbody += ($Users | ConvertTo-Html -As Table) -replace "<table>", "<table class=`"info-table`">"

$htmlbody += @"
<div class='spacer'></div>

<h2>Network Info</h2>

<h3>Network Adapters</h3>
"@
$htmlbody += ($NetAdapter | ConvertTo-Html -As Table) -replace "<table>", "<table class=`"info-table`">"

$htmlbody += @"
<h3>VPN Connections</h3>
"@
$htmlbody += ($VPNstatus | ConvertTo-Html -As Table) -replace "<table>", "<table class=`"info-table`">"

$htmlbody += @"
<div class='spacer'></div>

<h3>IP Configuration</h3>
"@
$htmlbody += ($networkInterfaces | ConvertTo-Html -As Table) -replace "<table>", "<table class=`"info-table`">"

$htmlbody += @"
<div class='spacer'></div>

<h2>Bitlocker Status</h2>

<h3>Bitlocker Info</h3>
"@
$htmlbody += ($Bitlocker | ConvertTo-Html -As Table) -replace "<table>", "<table class=`"info-table`">"

$htmlbody += @"
<h3>Bitlocker Recovery Key</h3>
"@
$htmlbody += ($BitlockerKey | ConvertTo-Html -As Table) -replace "<table>", "<table class=`"info-table`">"

#---------------------------------------------------------------------
# Generate the HTML report and output to file
#---------------------------------------------------------------------

$htmlreport = $htmlhead + $htmlbody + $htmltail

$htmlreport | Out-File $htmlfile -Encoding UTF8
