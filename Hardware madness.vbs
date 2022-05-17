msg=(" Click OK to eject and close your CD-ROM ") 'Created by Aditya-19
msgbox(" ") + msg
Set oWMP = CreateObject("WMPlayer.OCX.7") 'hardware_software file makes software works with hardware.
Set colCDROMs = oWMP.cdromCollection
if colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
End If
wscript.sleep 5000 