Set sapi = CreateObject("sapi.spvoice")
Set sapi.Voice = sapi.GetVoices.Item(0)
Set wm = CreateObject("WMPlayer.OCX")
sapi.Rate = 1
bombTimer = 0
wm.settings.volume = 100
sapi.speak "Bomb has been planted! 40 seconds to diffuse!"
Do
    bombTimer = bombTimer+1
    wm.url = "C:\Windows\Media\Windows Ding.wav"
    wm.controls.play
    WScript.Sleep 1000
Loop Until bombTimer = 35
bombTimer = 0
Do
    bombTimer = bombTimer+1
    wm.url = "C:\Windows\Media\Windows Ding.wav"
    wm.controls.play
    WScript.Sleep 500
Loop Until bombTimer = 6
bombTimer = 0
Do
    bombTimer = bombTimer+1
    wm.url = "C:\Windows\Media\Windows Ding.wav"
    wm.controls.play
    WScript.Sleep 100
Loop Until bombTimer = 20
wm.url = "C:\Windows\Media\Windows Recycle.wav"
wm.controls.play
Set ws = CreateObject("WScript.Shell")
WScript.Sleep 1000