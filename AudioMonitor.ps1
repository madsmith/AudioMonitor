# SoundVolumeView Binary
$SoundVolumeView = '.\SoundVolumeView.exe'

# This needs to be the name of the audio device used for game output
# as it appears in SoundVolumeView.
$TargetAudioDevice = "VB-Audio VoiceMeeter VAIO3"

# Load a JSON array of applications to monitor.
# Each entry must specify the application name and may specify an optional delay.
#   Name of application - exactly as it appeears in Task Manager.
#   Delay in Seconds - Generally, set it for longer than it takes the game
#                      to initialize it's sound engine.
#  Example:
#  [ {
#      "Name": "RSI Launcher",
#      "Delay": 2
#    }
#  ]

$jsonPath = Join-Path -Path $PSScriptRoot -ChildPath 'applications.json'

# Load the JSON file if it exists
if (Test-Path $jsonPath) {
    $applications = Get-Content $jsonPath | ConvertFrom-Json

    # Convert application data into a hashtable for runtime state tracking
    $runtimeState = @()
    foreach ($app in $applications) {
        $entry = @{
            Name = $app.Name
            Started = $false
        }
        if ("Delay" -in $app.psobject.Properties.Name) {
            $entry.Delay = $app.Delay
        }
        $runtimeState += $entry
    }
    $applications = $runtimeState
} else {
    Write-Host "No applications.json file found.  Please define a list of applications"
    Write-Host "to monitor in the following format and save it as applications.json"
    Write-Host "in the same directory as this script."
    Write-Host ""
    Write-Host "Example applications.json:"
    Write-Host "  [{"
    Write-Host "    ""Name"": ""RSI Launcher"","
    Write-Host "    ""Delay"": 2"
    Write-Host "  }]"
    exit
}

$DefaultDelay = 1
$ProcessPollInterval = 1

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

# Resolve the sound device
# Enumerate available sound devices
$soundDevices = Get-CimInstance -ClassName Win32_SoundDevice | Select-Object -ExpandProperty Name

# Check if the specified device is in the list of available devices
if ($TargetAudioDevice -notin $soundDevices) {
    Write-Host "The specified audio device '$TargetAudioDevice' was not found."
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
    $TargetAudioDevice = $soundDevices[$userSelection]
    Write-Host "You've selected the audio device: '$TargetAudioDevice'"
}

# Monitoring loop
while ($true) {
    $jobs = @()

    foreach ($app in $applications) {
        $process = Get-Process $app.Name -ErrorAction SilentlyContinue | Select-Object -First 1

        if ($process) {
            if (-not $app.Started) {
                $process_id = $process.Id

                if ($app.ContainsKey('Delay')) {
                    $delay = $app.Delay
                } else {
                    $delay = $DefaultDelay
                }

                $job = Start-Job -ScriptBlock {
                    param ($SoundVolumeView, $TargetAudioDevice, $appName, $process_id, $delay)

                    # Wait configured delay for program to attach to sound device
                    Start-Sleep -Seconds $delay

                    Write-Host "Fixing Audio Device for $appName"
                    & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" 0 "$process_id"
                } -ArgumentList $SoundVolumeView, $TargetAudioDevice, $app.Name, $process_id, $delay

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

    Start-Sleep -Seconds $ProcessPollInterval
}