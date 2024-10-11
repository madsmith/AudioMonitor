# Define new version
$newVersion = "1.4.0"

# Delete build directory
Remove-Item -Recurse -Force -Path "build/AudioMonitor"

# Delete old zip file if it exists
if (Test-Path "build/AudioMonitor-$newVersion.zip") {
  Remove-Item -Force -Path "build/AudioMonitor-$newVersion.zip"
}

# Create build Directory
New-Item -ItemType Directory -Path "build/AudioMonitor"

# Copy Files to build directory
Copy-Item -Path "AudioMonitor.ps1" -Destination "build/AudioMonitor/AudioMonitor.ps1"
Copy-Item -Path "AudioFixInteractive.ps1" -Destination "build/AudioMonitor/AudioFixInteractive.ps1"
Copy-Item -Path "AMLibModule.psm1" -Destination "build/AudioMonitor/AMLibModule.psm1"
Copy-Item -Path "sample_config.json" -Destination "build/AudioMonitor/sample_config.json"
Copy-Item -Path "Readme.md" -Destination "build/AudioMonitor/Readme.md"

# Create a new version of the zip file
Compress-Archive -Path "build/AudioMonitor" -DestinationPath "build/AudioMonitor-$newVersion.zip"

