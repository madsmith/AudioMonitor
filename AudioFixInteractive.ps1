# SoundVolumeView Binary
$SoundVolumeView = '.\SoundVolumeView.exe'

# This needs to be the name of the audio device used for game output
# as it appears in SoundVolumeView.
$TargetAudioDevice = "VB-Audio VoiceMeeter VAIO3"

# End Configuration

$targetApplication = $args[0]

# if $args[1] is set, use it as the audio device
if ($args[1]) {
  $TargetAudioDevice = $args[1]
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

# Enumerate Processes
$processes = Get-Process | Select-Object -ExpandProperty ProcessName

# Temporary file to store output, placed in the Windows temp folder
$tempFile = Join-Path -Path $env:TEMP -ChildPath "temp_audio_sessions.csv"

# Run SoundVolumeView to dump audio sessions to the temporary file
& $SoundVolumeView /scomma $tempFile

# Wait for the file to be written
Start-Sleep -Seconds 1

# Read and process the file
if (Test-Path $tempFile) {
  $audioSessions = Import-Csv $tempFile

  # Build unique list of applications
  $uniqueApplications = @{}
  foreach ($session in $audioSessions) {
      #Write-Host "Session: $($session.'Name')"
      $appName = $session.'Name'
      if ($processes -contains $appName) {
          $uniqueApplications[$appName] = $true
      }
  }
  # Convert hashtable keys to an array for easier manipulation
  $appNames = @($uniqueApplications.Keys)

  if (-not ($targetApplication -and $appNames -contains $targetApplication)) {
    # List unique applications
    Write-Host "Select an application to fix or 0 to exit:"
    Write-Host "  1: Target Focused Application (3s delay)"
    for ($i = 0; $i -lt $appNames.Count; $i++) {
        Write-Host "  $($i+2): $($appNames[$i])"
    }

    # Prompt the user to select an application
    [int]$userSelection = Read-Host "Enter the number corresponding to your desired application"
    while ($userSelection -lt 0 -or $userSelection -ge ($appNames.Count + 2)) {
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
      $targetApplication = $appNames[$userSelection - 2]
      Write-Host "You've selected the application: '$targetApplication'"
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
  while ($countDown -gt 0) {
    Write-Host "Targeting foreground application in $countDown seconds..."
    Start-Sleep -Seconds 1
    $countDown--
  }
  Write-Host "Setting audio device for focused application to '$TargetAudioDevice'"
  & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" 0 "focused"
} else {
  Write-Host "Setting audio device for '$targetApplication' to '$TargetAudioDevice'"
  $binaryPath = (Get-Process $targetApplication | Select-Object -First 1).MainModule.FileName
  #Write-Host "Binary Path: $binaryPath"
  & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" 0 "$binaryPath"
}