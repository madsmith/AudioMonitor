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