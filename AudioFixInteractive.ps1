# SoundVolumeView Binary
$SoundVolumeView = '.\SoundVolumeView.exe'

# This needs to be the name of the audio device used for game output
# as it appears in SoundVolumeView.
$TargetAudioDevice = "Voicemeeter VAIO3 Input"

# End Configuration

$DeviceType = 0
$positional = 0
for ($i = 0; $i -lt $args.Count; $i++) {
  if ($args[$i] -eq "-h" -or $args[$i] -eq "--help") {
    Write-Host "Usage: AudioFixInteractive.ps1 [--type <0|1|2|all] [application] [audio device]"
    Write-Host "  application: (optional) The name of the application to fix audio for."
    Write-Host "  audio device: (optional) The name of the audio device to use."
    Write-Host "  --type <0|1|2|all>: (optional) The type of audio device to fix."
    Write-Host "      0: Console, 1: Multimedia, 2: Communications, all: All types."
    Write-Host "If no application is specified, a list of applications will be displayed."
    Write-Host "If no audio device is specified, a list of audio devices will be displayed."
    exit
  }

  if ($args[$i] -eq "--type") {
    if (@("0", "1", "2", "all") -contains $args[$i+1]) {
      $DeviceType = $args[$i+1]
      $i++
      continue
    } else {
      Write-Host "Invalid device type specified.  Please specify 0, 1, 2, or all."
      exit
    }
  }

  if ($positional -eq 0) {
    $targetApplication = $args[$i]
    $positional++
  } elseif ($positional -eq 1) {
    $TargetAudioDevice = $args[$i]
    $positional++
  }
}

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

# Temporary file to store output, placed in the Windows temp folder
$tempFile = Join-Path -Path $env:TEMP -ChildPath "temp_audio_sessions.csv"

# Run SoundVolumeView to dump audio sessions to the temporary file
& $SoundVolumeView /scomma $tempFile /Columns "Name,Type,Process ID"

# Wait for the file to be written
Start-Sleep -Seconds 1

# Read and process the file
if (Test-Path $tempFile) {
  $audioSessions = Import-Csv $tempFile

  # Build unique list of applications
  $uniqueApplications = @{}
  $foundTargetApplication = $false
  foreach ($session in $audioSessions) {
    #Write-Host "Session: $($session.'Name')"
    if ($session.'Type' -ne 'Application') {
        continue
    }
    $appName = $session.'Name'
    $processID = $session.'Process ID'
    $uniqueApplications[$processID] = @{
      'Name' = $appName
      'PID' = $processID
    }
    if ($targetApplication -and $appName -eq $targetApplication) {
      $foundTargetApplication = $true
    }
  }

  # List out applications in alphabetical order by name
  $applications = @($uniqueApplications.Values) | ForEach-Object { [PSCustomObject]$_ } | Sort-Object -Property Name

  if (-not $foundTargetApplication) {
    # List unique applications
    Write-Host "Select an application to fix or 0 to exit:"
    Write-Host "  1: Target Focused Application (3s delay)"
    for ($i = 0; $i -lt $applications.Count; $i++) {
        Write-Host "  $($i+2): $($applications[$i].Name)"
    }

    # Prompt the user to select an application
    [int]$userSelection = Read-Host "Enter the number corresponding to your desired application"
    while ($userSelection -lt 0 -or $userSelection -ge ($applications.Count + 2)) {
        Write-Host "Invalid selection."
        $userSelection = Read-Host "Please enter a valid number corresponding to your desired application"
    }

    if ($userSelection -eq 0) {
      exit
    }
    if ($userSelection -eq 1) {
      $targetApplication = "::focused::"
      Write-Host "Targetting current focused application with 3s delay."
    } else {
      # Retrieve the selected application name
      $targetApplication = $applications[$userSelection - 2].Name
      Write-Host "You've selected the application: '$($targetApplication)'"
    }
  }

  Remove-Item $tempFile -Force # Clean up
} else {
  Write-Host "Failed to retrieve audio sessions."
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
        Write-Host "  $($i + 1): $($soundDevices[$i])"
    }

    # Prompt the user to select a device
    [int]$userSelection = Read-Host "Enter the number corresponding to your desired audio device"
    while ($userSelection -lt 1 -or $userSelection -ge ($soundDevices.Count + 1)) {
        Write-Host "Invalid selection."
        [int]$userSelection = Read-Host "Please enter a valid number corresponding to your desired audio device"
    }

    # Update the target audio device based on user selection
    $TargetAudioDevice = $soundDevices[$userSelection - 1]
    Write-Host "You've selected the audio device: '$TargetAudioDevice'"
}

if ($targetApplication -eq "::focused::") {
  $countDown = 3
  Write-Host "Targeting focused application in 3 seconds..."
  while ($countDown -gt 1) {
    Start-Sleep -Seconds 1
    $countDown--
    Write-Host "  $countDown seconds..."
  }
  Write-Host "Setting audio device for focused application to '$TargetAudioDevice'"
  & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" $DeviceType "focused"
} else {
  $appName = $targetApplication
  Write-Host "Setting audio device for '$appName' to '$TargetAudioDevice'"
  foreach ($process in $uniqueApplications.Values) {
    if ($process.Name -eq $targetApplication) {
      & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" $DeviceType $process.PID
    }
  }
}