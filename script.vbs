Set objShell = CreateObject ("WScript.Shell")

Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

Do
	' Play audio
	oPlayer.URL = "C:\Windows\Media\Windows Critical Stop.wav"
	oPlayer.controls.play

	' Display a message box with the title "Hello" and the text "This is a popup with sound"
	objShell.Popup "Are you being productive?" & vbCrLf & "If yes, congratulations!" & vbCrLf & "If no, get to work loser!", infinite, "Are you being productive?", 36

	WScript.Sleep 300000
Loop
