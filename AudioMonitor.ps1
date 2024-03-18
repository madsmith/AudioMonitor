# SoundVolumeView Binary
$SoundVolumeView = '.\SoundVolumeView.exe'

$DefaultDelay = 1
$sampleConfig = @'
{
    "targetAudioDevice": "Voicemeeter VAIO3 Input",
    "processPollingInterval": 1,
    "monitoredApplications": [
        { "Name": "RSI Launcher", "Delay": 2 },
        { "Name": "StarCitizen", "Delay": 20 },
        { "Name": "Helldivers2", "Delay": 37 },
        { "Name": "CrabChampions-Win64-Shipping"}
    ]
}
'@

# Path to the config.json file
$jsonPath = Join-Path -Path $PSScriptRoot -ChildPath 'config.json'

function Assert-SoundVolumeView {
    # Resolve Sound Volume View Path
    try {
        # Attempt to resolve the path to SoundVolumeView
        $resolvedPath = Resolve-Path $SoundVolumeView -ErrorAction Stop
        $script:SoundVolumeView = $resolvedPath.ProviderPath
    } catch {
        # If the path cannot be resolved, inform the user and exit the script
        Write-Host "Unable to find SoundVolumeView.exe at the specified path: $SoundVolumeView"
        Write-Host "Please ensure SoundVolumeView.exe exists at the correct path and try again."
        Write-Host "If SoundVolumeView is not installed, please download and install it."
        Write-Host "Alternatively, edit the script to specify the correct path to SoundVolumeView.exe."
        exit
    }
}

function Get-SVVTargets {
    Assert-SoundVolumeView

    # Temporary file to store output, placed in the Windows temp folder
    $tempFile = Join-Path -Path $env:TEMP -ChildPath "temp_audio_sessions.csv"
    & $SoundVolumeView /scomma $tempFile /Columns "Name,Type,Direction,Device Name,Process ID,Window Title,Process Path" | Out-Null

    $svvData = Import-Csv -Path $tempFile
    Remove-Item -Path $tempFile

    $soundDevices = $svvData |
        Where-Object { $_.'Type' -eq 'Device' -and $_.'Direction' -eq 'Render' } |
        Select-Object -ExpandProperty 'Name' |
        Sort-Object

    # Saving this data for future uses.  It's more accurate than relying on just the list of running
    # processes since it's just active audio processes, but I'm not using it yet for the monitor loop
    # for concerns about the performance of running SoundVolumeView every poll vs just using Get-Process
    $applications = $svvData |
        Where-Object { $_.'Type' -eq 'Application' } |
        Group-Object -Property 'Process ID' |
        ForEach-Object {
            $_.Group | Select-Object -First 1
        } |
        ForEach-Object {
            [PSCustomObject]@{
                Name = $_.'Name'
                ProcessID = $_.'Process ID'
                WindowTitle = $_.'Window Title'
                ProcessPath = $_.'Process Path'
            }
        } |
        Sort-Object -Property 'Name'

    $result = New-Object PSObject -Property @{
        SoundDevices = $soundDevices
        Applications = $applications
    }

    return $result
}

function Confirm-ResetConfig {
    Write-Host "Would you like to reset it to the default configuration? (y/n)"
    $response = Read-Host
    if ($response -eq "y") {
        Write-Host "Saving a backup of the existing config.json file to config.json.bak"
        $backupPath = Join-Path -Path $PSScriptRoot -ChildPath 'config.json.bak'
        Write-Host "Backing up the existing config.json file to config.json.bak"
        Copy-Item -Path $jsonPath -Destination $backupPath
        Write-SampleConfig
    } else {
        exit
    }
}
function Write-SampleConfig {
    Write-Host "Generating default config.json file.  Please edit config.json to your needs."
    Write-Host $sampleConfig
    Set-Content -Path $jsonPath -Value $sampleConfig

    $newConfig = $sampleConfig | ConvertFrom-Json
    $script:config = $newConfig
}

# Function to prompt for the user to select an audio device
function Select-AudioDevice {
    Write-Host "Please select a valid audio device from the following list:"

    # Display the list of devices with index
    for ($i = 0; $i -lt $soundDevices.Count; $i++) {
        Write-Host "${i}: $($soundDevices[$i])"
    }

    # Prompt the user to select a device
    [int]$userSelection = Read-Host "Enter the number corresponding to your desired audio device"
    while ($userSelection -lt 0 -or $userSelection -ge $soundDevices.Count) {
        [int]$userSelection = Read-Host "Invalid selection. Please enter a valid number corresponding to your desired audio device"
    }

    # Update the target audio device based on user selection
    $script:config.TargetAudioDevice = $soundDevices[$userSelection]
    Write-Host "You've selected the audio device: '$($config.TargetAudioDevice)'."
    $config | ConvertTo-Json -Depth 5 | Set-Content -Path $jsonPath
}

function DecodeDeviceType {
    param (
        [object]$deviceType
    )

    if ($deviceType -is [string]) {
        switch ($deviceType.ToLower()) {
            "console" { return 0 }
            "multimedia" { return 1 }
            "communications" { return 2 }
            "all" { return "all" }
            default {
                Write-Host "Invalid DeviceType: $deviceType"
                exit
            }
        }
    }

    if ($deviceType -is [int]) {
        if ($deviceType -ge 0 -and $deviceType -le 2) {
            return $deviceType
        } else {
            Write-Host "Invalid DeviceType: $deviceType"
            exit
        }
    }
}

# If json path is missing, show usage
if (-not (Test-Path $jsonPath)) {
    Write-Host "No config.json file found."
    Write-SampleConfig
    exit
}

# Load the config file
$config = Get-Content $jsonPath | ConvertFrom-Json

if (-not $config) {
    Write-Host "The config.json file is not valid JSON."
    Confirm-ResetConfig
}

$targets = Get-SVVTargets
$soundDevices = $targets.SoundDevices

# Validate the config.json file
if (-not $config.psobject.Properties.Name.Contains('processPollingInterval')) {
    Write-Host "The config.json file is missing the 'processPollingInterval' property."
    Add-Member -InputObject $config -MemberType NoteProperty -Name "processPollingInterval" -Value 1
    $config | ConvertTo-Json -Depth 5 | Set-Content -Path $jsonPath
}

if (-not $config.psobject.Properties.Name.Contains('targetAudioDevice')) {
    Write-Host "The config.json file is missing the 'targetAudioDevice' property."
    Add-Member -InputObject $config -MemberType NoteProperty -Name "targetAudioDevice" -Value ""
    Select-AudioDevice
}

if (-not ($config.targetAudioDevice -in $soundDevices)) {
    Write-Host "The specified audio device '$($config.targetAudioDevice)' was not found."
    Select-AudioDevice
}

# Convert application data into a hashtable for runtime state tracking
$runtimeState = @()
foreach ($app in $config.monitoredApplications) {
    $entry = @{
        Name = $app.Name
        Started = $false
    }
    if ("Delay" -in $app.psobject.Properties.Name) {
        $entry.Delay = $app.Delay
    } else {
        $entry.Delay = $DefaultDelay
    }
    if ("DeviceType" -in $app.psobject.Properties.Name) {
        $entry.DeviceType = DecodeDeviceType $app.DeviceType
    } else {
        $entry.DeviceType = 0
    }
    $runtimeState += $entry
}
$applications = $runtimeState

# Resolve Sound Volume View Path
Assert-SoundVolumeView

# Monitoring loop
while ($true) {
    $jobs = @()

    foreach ($app in $applications) {
        $processes = Get-Process $app.Name -ErrorAction SilentlyContinue

        if ($processes -and $processes.Count -gt 0) {
            $multiple_processes = $processes.Count -gt 1
            if (-not $app.Started) {
                if ($multiple_processes) {
                    $pids = $processes | Select-Object -ExpandProperty Id
                } else {
                    $pids = $processes | Select-Object -First 1 -ExpandProperty Id
                }

                $delay = $app.Delay

                $job = Start-Job -ScriptBlock {
                    param ($SoundVolumeView, $TargetAudioDevice, $appName, $DeviceType, $process_list, $delay, $multiple)

                    # Wait configured delay for program to attach to sound device
                    Start-Sleep -Seconds $delay

                    if ($multiple) {
                        $pid_list = $process_list -join ', '
                        Write-Host "Fixing Audio Device for $appName [pids: $pid_list]"
                    } else {
                        Write-Host "Fixing Audio Device for $appName"
                    }

                    foreach ($process_id in $process_list) {
                        # Update console type audio output
                        & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" $DeviceType "$process_id"
                    }
                } -ArgumentList $SoundVolumeView, $config.TargetAudioDevice, $app.Name, $app.DeviceType, $pids, $delay, $multiple_processes

                $jobs += $job
                $app.Started = $true
            }
        } else {
            $app.Started = $false
        }
    }

    # Wait for all jobs to complete
    $jobs | Wait-Job | Receive-Job

    # Clean up completed jobs
    $jobs | Remove-Job

    Start-Sleep -Seconds $config.processPollingInterval
}