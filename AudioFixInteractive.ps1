Import-Module ".\AMLibModule.psm1" -Force

$config = Get-Config

$TargetApplication = $null
$TargetAudioDevice = $null

$DelaySeconds = 5

#if ($config.targetAudioDevice) {
#  $targetAudioDevice = $config.targetAudioDevice
#}

#####################################################################
# Argument Parsing
#####################################################################

$DeviceType = 0
$position = 0
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

  if ($position -eq 0) {
    $TargetApplication = $args[$i]
    $position++
  } elseif ($position -eq 1) {
    $TargetAudioDevice = $args[$i]
    $position++
  }
}

#####################################################################
# Begin Script Execution
#####################################################################

$SoundVolumeView = Get-SoundVolumeView

$targets = Get-SoundVolumeViewTargets
$applications = $targets.Applications

# Is there a valid target application?
$foundTargetApplication = $false
if ($TargetApplication) {
  foreach ($app in $applications) {
    if ($app.Name -eq $TargetApplication) {
      $foundTargetApplication = $true
      break
    }
  }
}

# Prompt for an application to target
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
    $TargetApplication = "::focused::"
    Write-Host "Targetting current focused application with $($DelaySeconds)s delay..."
  } else {
    # Retrieve the selected application name
    $TargetApplication = $applications[$userSelection - 2].Name
    Write-Host "You've selected the application: '$($TargetApplication)'"
  }
}

$soundDevices = $target.SoundDevices

# Check if the specified device is in the list of available devices
if ($null -eq $TargetAudioDevice -or $TargetAudioDevice -notin $soundDevices) {
  if ($null -ne $TargetAudioDevice) {
    Write-Host "The specified audio device '$TargetAudioDevice' was not found."
  }
  $TargetAudioDevice = Select-AudioDevice -UpdateConfig $false
}

if ($TargetApplication -eq "::focused::") {
  $countDown = $DelaySeconds
  Write-Host "Targeting focused application in $DelaySeconds seconds..."
  while ($countDown -gt 1) {
    Start-Sleep -Seconds 1
    $countDown--
    Write-Host "  $countDown seconds..."
  }
  Write-Host "Setting audio device for focused application to '$TargetAudioDevice'"
  & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" $DeviceType "focused"
} else {
  $appName = $TargetApplication
  Write-Host "Setting audio device for '$appName' to '$TargetAudioDevice'"
  foreach ($process in $uniqueApplications.Values) {
    if ($process.Name -eq $TargetApplication) {
      & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" $DeviceType $process.PID
    }
  }
}