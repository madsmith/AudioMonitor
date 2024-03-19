Import-Module ".\AMLibModule.psm1" -Force

$DefaultDelay = 1

# Load Config
$config = Get-Config
$SoundVolumeView = Get-SoundVolumeView

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