# Define new version
$newVersion = "1.2.0"

# Delete build directory
Remove-Item -Recurse -Force -Path "build/AudioMonitor"

# Create build Directory
New-Item -ItemType Directory -Path "build/AudioMonitor"

# Copy Files to build directory
Copy-Item -Path "AudioMonitor.ps1" -Destination "build/AudioMonitor/AudioMonitor.ps1"
Copy-Item -Path "AudioFixInteractive.ps1" -Destination "build/AudioMonitor/AudioFixInteractive.ps1"
Copy-Item -Path "sample_config.json" -Destination "build/AudioMonitor/sample_config.json"
Copy-Item -Path "Readme.md" -Destination "build/AudioMonitor/Readme.md"

# Create a new version of the zip file
Compress-Archive -Path "build/AudioMonitor" -DestinationPath "build/AudioMonitor-$newVersion.zip"

