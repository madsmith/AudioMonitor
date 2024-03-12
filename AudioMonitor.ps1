# SoundVolumeView Binary
$SoundVolumeView = '.\SoundVolumeView.exe'

$DefaultDelay = 1
$sampleConfig = @'
{
    "targetAudioDevice": "VB-Audio VoiceMeeter VAIO3",
    "processPollingInterval": 1,
    "applications": [
        { "Name": "RSI Launcher", "Delay": 2 },
        { "Name": "StarCitizen", "Delay": 20 },
        { "Name": "Helldivers", "Delay": 37 },
        { "Name": "CrabChampions"}
    ]
}
'@

# Path to the config.json file
$jsonPath = Join-Path -Path $PSScriptRoot -ChildPath 'config.json'

# Enumerate available sound devices
$soundDevices = Get-CimInstance -ClassName Win32_SoundDevice | Select-Object -ExpandProperty Name

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
    $config.TargetAudioDevice = $soundDevices[$userSelection]
    Write-Host "You've selected the audio device: '$($config.TargetAudioDevice)'."
    $config | ConvertTo-Json -Depth 5 | Set-Content -Path $jsonPath
    exit
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

$applications = $config.applications

# Convert application data into a hashtable for runtime state tracking
$runtimeState = @()
foreach ($app in $applications) {
    $entry = @{
        Name = $app.Name
        Started = $false
    }
    if ("Delay" -in $app.psobject.Properties.Name) {
        $entry.Delay = $app.Delay
    } else {
        $entry.Delay = $DefaultDelay
    }
    $runtimeState += $entry
}
$applications = $runtimeState

# Resolve Sound Volume View Path
try {
    # Attempt to resolve the path to SoundVolumeView
    $resolvedPath = Resolve-Path $SoundVolumeView -ErrorAction Stop
    $SoundVolumeView = $resolvedPath.ProviderPath
} catch {
    # If the path cannot be resolved, inform the user and exit the script
    Write-Host "Unable to find SoundVolumeView.exe at the specified path: $SoundVolumeView"
    Write-Host "Please ensure SoundVolumeView.exe exists at the correct path and try again."
    Write-Host "If SoundVolumeView is not installed, please download and install it."
    Write-Host "Alternatively, edit the script to specify the correct path to SoundVolumeView.exe."
    exit
}

# Monitoring loop
while ($true) {
    $jobs = @()

    foreach ($app in $applications) {
        $processes = Get-Process $app.Name -ErrorAction SilentlyContinue | Select-Object -First 1

        if ($processes -and $processes.Count -gt 0) {
            foreach ($process in $processes) {
                if (-not $app.Started) {
                    $process_id = $process.Id

                    $delay = $app.Delay

                    $job = Start-Job -ScriptBlock {
                        param ($SoundVolumeView, $TargetAudioDevice, $appName, $process_id, $delay)

                        # Wait configured delay for program to attach to sound device
                        Start-Sleep -Seconds $delay

                        Write-Host "Fixing Audio Device for $appName [id: $process_id]"
                        & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" 0 "$process_id"
                    } -ArgumentList $SoundVolumeView, $config.TargetAudioDevice, $app.Name, $process_id, $delay

                    $jobs += $job
                    $app.Started = $true
                }
            }
            $app.Started = $true
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