Set oWMP = CreateObject("WMPlayer.OCX" )
Set colCDROMs = oWMP.cdromCollection
while true	
	colCDROMs.Item(i).Eject
	Wscript.Sleep 300
wend