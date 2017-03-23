# Eventghost Skype Plugin
This is a basic plugin for the automation program [Eventghost](http://www.eventghost.org/) that provides events to reflect the lifecycle of Skype calls. At present, it supplies the following events:

* Skype.Calling
    _Occurs when a call attempt begins ringing_

* Skype.Refused
    _Occurs when the person called does not accept the call_

* Skype.Cancelled
    _Occurs when the person calling cancels the call before it is picked up_

* Skype.Call in Progress
    _Occurs when the call is answered and connected_

* Skype.Finished
    _Occurs when either party hangs up a connected call_

Other events are possible - the event names are the result of the `Skype4COM.Skype.Convert.CallStatusToText` function, and issued whenever that status changes.

# Installing
#### 32-Bit Windows
1. Download the latest release from the [releases](https://github.com/OrbitalDan/EventghostSkypePlugin/releases) page.
2. Extract the Skype folder into the plugins folder of Eventghost. Typically, this will be `C:\Program Files\Eventghost\plugins`.
3. Copy Skype's COM API DLL (`Skype4COM.dll`) from `C:\Program Files (x86)\Common Files\Skype` to the system folder `C:\Windows\System32`
4. From an elevated CMD prompt (search start menu for CMD.exe, right-click, 'Run as Administrator...'), navigate to `C:\Windows\System32` and run the following command:
```bat
C:\Windows\System32> regsvr32.exe Skype4COM.dll
```
5. The first time the plugin is loaded, Eventghost may freeze while Skype displays a prompt asking whether or not to allow Eventghost.exe access.  You'll need to accept/decline before Eventghost will continue running.

#### 64-Bit Windows
1. Download the latest release from the [releases](https://github.com/OrbitalDan/EventghostSkypePlugin/releases) page.
2. Extract the Skype folder into the plugins folder of Eventghost. Typically, this will be `C:\Program Files (x86)\Eventghost\plugins`.
3. Copy Skype's COM API DLL (`Skype4COM.dll`) from `C:\Program Files (x86)\Common Files\Skype` to the 32-bit system folder `C:\Windows\SysWOW64`
4. From an elevated CMD prompt (search start menu for CMD.exe, right-click, 'Run as Administrator...'), navigate to `C:\Windows\SysWOW64` and run the following command:
```bat
C:\Windows\SysWOW64> regsvr32.exe Skype4COM.dll
```
5. The first time the plugin is loaded, Eventghost may freeze while Skype displays a prompt asking whether or not to allow Eventghost.exe access.  You'll need to accept/decline before Eventghost will continue running.
