$config = $null

# Path to the config.json file
$configPath = Join-Path -Path $PSScriptRoot -ChildPath 'config.json'

$sampleConfig = @'
{
    "SoundVolumeView": "./SoundVolumeView.exe",
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

# Sound Devices
$SoundDevices = $null

#####################################################################
# Config Functions
#####################################################################

function Import-Config {
  # If json path is missing, show usage
  if (-not (Test-Path $script:configPath)) {
      Write-Host "No config.json file found."
      Write-SampleConfig
      exit
  }

  $config = Get-Content $script:configPath | ConvertFrom-Json

  # Check if we failed to load the config
  if (-not $config) {
    Write-Host "The config.json file is not valid JSON."
    Confirm-ResetConfig
  } else {
    $script:config = $config
  }

  # Validate the config.json file
  if (-not $script:config.psobject.Properties.Name.Contains('processPollingInterval')) {
    Write-Host "The config.json file is missing the 'processPollingInterval' property."
    Add-Member -InputObject $script:config -MemberType NoteProperty -Name "processPollingInterval" -Value 1
    $script:config | ConvertTo-Json -Depth 5 | Set-Content -Path $script:configPath
  }

  if (-not $script:config.psobject.Properties.Name.Contains('targetAudioDevice')) {
    Write-Host "The config.json file is missing the 'targetAudioDevice' property."
    Add-Member -InputObject $script:config -MemberType NoteProperty -Name "targetAudioDevice" -Value ""
    Select-AudioDevice
  }
}

function Get-Config {
  if ($null -eq $script:config) {
    Import-Config

    if ($null -eq $script:config) {
      Write-Host "Error: Config is null."
    }
  }

  return $script:config
}

function Write-DefaultConfig {
  Write-Host "Generating default config.json file.  Please edit config.json to your needs."
  Write-Host $script:sampleConfig
  Set-Content -Path $jsonPath -Value $script:sampleConfig

  $newConfig = $script:sampleConfig | ConvertFrom-Json
  # Update the config variable
  $script:config = $newConfig
}

function Confirm-ResetConfig {
  Write-Host "Would you like to reset config to the default configuration? (y/n)"
  $response = Read-Host
  if ($response -eq "y") {
      Write-Host "Saving a backup of the existing config.json file to config.json.bak"
      $backupPath = Join-Path -Path $PSScriptRoot -ChildPath 'config.json.bak'
      Write-Host "Backing up the existing config.json file to config.json.bak"
      Copy-Item -Path $configPath -Destination $backupPath
      Write-SampleConfig
  } else {
      exit
  }
}

#####################################################################
# Sound Volume View Accessors
#####################################################################

function Get-SoundVolumeView {
  $svvPath = $script:config.SoundVolumeView
  if (-not $svvPath) {
    $svvPath = ".\SoundVolumeView.exe"
  }

  # Resolve Sound Volume View Path
  try {
    # Attempt to resolve the path to SoundVolumeView
    $resolvedPath = Resolve-Path $svvPath -ErrorAction Stop
    return $resolvedPath.ProviderPath
  } catch {
    # If the path cannot be resolved, inform the user and exit the script
    Write-Host "Unable to find SoundVolumeView.exe at the specified path: $svvPath"
    Write-Host "Please ensure SoundVolumeView.exe exists at the correct path and try again."
    Write-Host "If SoundVolumeView is not installed, please download and install it."
    Write-Host "Alternatively, edit the script to specify the correct path to SoundVolumeView.exe."
    exit
  }
}

function Get-SoundDevices {
  if ($null -eq $script:SoundDevices) {
    $targets = Get-SoundVolumeViewTargets
    $script:SoundDevices = $targets.SoundDevices
  }
  return $script:SoundDevices
}

function Get-SoundVolumeViewTargets {
  $SoundVolumeView = Get-SoundVolumeView

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

#####################################################################
# Interactive Prompts
#####################################################################

function Get-TargetAudioDevice {
  $soundDevices = Get-SoundDevices

  if (-not ($script:config.targetAudioDevice -in $soundDevices)) {
    Write-Host "The specified audio device '$($script:config.targetAudioDevice)' was not found."
    Select-AudioDevice
  }
}

# Function to prompt for the user to select an audio device
function Select-AudioDevice {
  param (
    [bool]$UpdateConfig = $true
  )

  $soundDevices = Get-SoundDevices

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

  $TargetAudioDevice = $soundDevices[$userSelection]

  # Update the target audio device based on user selection
  Write-Host "You've selected the audio device: '$TargetAudioDevice'."

  if ($UpdateConfig) {
    $script:config.TargetAudioDevice = $TargetAudioDevice
    $script:config | ConvertTo-Json -Depth 5 | Set-Content -Path $script:configPath
  }

  return $TargetAudioDevice
}



#####################################################################
# Utility Functions
#####################################################################

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