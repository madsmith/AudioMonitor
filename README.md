# Installation Steps
Download [SoundVolumeView by NirSoft](https://www.nirsoft.net/utils/sound_volume_view.html), link is towards the bottom of this page.  I grabbed the 64bit variant.

Download the two powershell scripts from this project.

Extract SoundVolumeView.exe into the folder with the powershell scripts.

Run the script in powershell.  e.g. Right click on script and select *"Run with Powershell"*

# AudioFixInteractive

This is a simple script which lists processes known by SoundVolumeView and applies an audio device fix to one process. It will interactively prompt you to specify which process or sound device to use.  A default sound device can be preset in the script my modifying **TargetAudioDevice**.

## Usage

    AudioFixInteractive.ps1 [<Application> [<AudioDevice>]]

If you specify an application name on the command line, it will attempt to find and adjust that process.

If you specify an application name and an audio device it will force that application to use that specify audio device.

Otherwise the script will prompt you interactively.

    C:\Users\Martin\Projects\AudioMonitor> .\AudioFixInteractive.ps1
    Select an application to fix or 0 to exit:
      1: Steam
      2: Spotify
      3: NVIDIA Broadcast
      4: Firefox
      5: RSI Launcher
      6: Discord
    Enter the number corresponding to your desired application: 5
    You've selected the application: 'RSI Launcher'
    The specified audio device 'UserUnsetAudioDevice' was not found.
    Please select a valid audio device from the following list:
      1: NVIDIA High Definition Audio
      2: USB Audio Device
      3: USB Audio Device
      4: Steam Streaming Microphone
      5: USB Audio Device
      6: Virtual Audio Cable
      7: Steam Streaming Speakers
      8: NVIDIA Virtual Audio Device (Wave Extensible) (WDM)
      9: NVIDIA Broadcast
      10: USB Audio Device
      11: High Definition Audio Device
      12: VB-Audio VoiceMeeter VAIO
      13: VB-Audio VoiceMeeter AUX VAIO
      14: VB-Audio VoiceMeeter VAIO3
      15: VB-Audio Virtual Cable
      16: VB-Audio Hi-Fi Cable
    Enter the number corresponding to your desired audio device: 14
    You've selected the audio device: 'VB-Audio VoiceMeeter VAIO3'
    Setting audio device for 'RSI Launcher' to 'VB-Audio VoiceMeeter VAIO3'


# AudioMonitor
 Monitor running processes and adjust sound device by invoking SoundVolumeView.exe.  SoundVolumeView is a toolkit for manipulating program audio devices.  It has a desktop mode which audio details about running processes and sound devices but it also features extensive command line options to tweaking audio. This tool uses the following command to adjust the audio renderer for various processes as it detects them running.

    SoundVolumeView.exe /SetAppDefault "VB-Audio VoiceMeeter VAIO3" 0 program.exe

 This script may require minor edits to run correctly.  It assumes that the program SoundVolumeView is available.  The default configuration is that this program should be in the same directory as the powershell script.

 [SoundVolumeView by NirSoft](https://www.nirsoft.net/utils/sound_volume_view.html)

 The powershell script has a setting for the target audio device.  The default is **"VB-Audio VoiceMeeter VAIO3"** as the author uses [VoiceMeeter Potato](https://vb-audio.com/Voicemeeter/potato.htm), but you may adjust the script for your target audio device.  To do so, edit the script to update **$TargetAudioDevice** to the name of your audio device as it appears when running SoundVolumeView.

     $TargetAudioDevice = "VB-Audio VoiceMeeter VAIO3"

 When run, the script will monitor the system for a list of applications as defined in the script.  When discovering that application, the script will use SoundVolumeView to adjust that applications console audio device to the target audio device.

 This changes the output device for that listed executable to "VB-Audio VoiceMeeter VAIO3".

 The application is then remembered as "started" and the script will not attempt to adjust that application again until it no longer sees that application as started.

 The applications that are monitored are in a list of applications in the script.  You may customize it for your purposes.  There are three parameters that need to be specified for each application.

  1. **Application Name** - [**Required**] - The application name as it appears in task manager must be specified.
  2. **Delay** - [*Optional*] - Sound device adjustment must occur after the executable has attached to a sound device, use the configurable delay (in seconds) to allow the monitored program time to initialize its sound.

  Here is an example application list.  Each entry must specify Name, can specify Delay and each entry should be seperated by a ','.

    $applications = @(
      @{
          Name = "RSI Launcher"
          Delay = 2
      },
      @{
          Name = "StarCitizen"
          Delay = 20
      },
      @{
          Name = "valheim"
          Delay = 14
      }
    )

