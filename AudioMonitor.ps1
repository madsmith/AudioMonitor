# Define an array of application information

# SoundVolumeView Binary
$SoundVolumeView = 'C:\Program Files\SoundVolumeView\SoundVolumeView.exe'

# This needs to be the name of the audio device used for game output
# as it appears in SoundVolumeView.
$TargetAudioDevice = "VB-Audio VoiceMeeter VAIO3"

# List of applications to monitor.
# Must specify 3 things
#   Name of application - exactly as it appeears in Task Manager.
#   Name of executable for said application [GameBinary.exe]
#   Delay in Seconds - Generally, set it for longer than it takes the game
#                      to initialize it's sound engine.
$applications = @(
    @{
        Name = "RSI Launcher"
        Executable = "RSI Launcher.exe"
        Delay = 2
    },
    @{
        Name = "StarCitizen"
        Executable = "StarCitizen.exe"
        Delay = 20
    },
    @{
        Name = "valheim"
        Executable = "valheim.exe"
        Delay = 14
    },
    @{
        Name = "eqgame"
        Executable = "eqgame.exe"
        Delay = 4
    },
    @{
        Name = "Roboquest"
        Executable = "RoboQuest-Win64-Shipping.exe"
        Delay = 4
    }
)

while ($true) {
    $jobs = @()

    foreach ($app in $applications) {
        $process = Get-Process $app.Name -ErrorAction SilentlyContinue

        if ($process) {
            if (-not $app.Started) {
                $job = Start-Job -ScriptBlock {
                    param ($SoundVolumeView, $TargetAudioDevice, $appName, $delay, $executable)
                    Start-Sleep -Seconds $delay
                    Write-Host "Fixing Audio Device for $appName"
                    & $SoundVolumeView /SetAppDefault "$TargetAudioDevice" 0 $executable
                } -ArgumentList $SoundVolumeView, $TargetAudioDevice, $app.Name, $app.Delay, $app.Executable

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

    Start-Sleep -Seconds 1
}