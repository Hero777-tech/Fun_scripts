'Created by Aditya-18
set wshshell = wscript.CreateObject("wScript.Shell")

Dim x

x=1

Do While x<5

wshshell.run("Notepad.exe")

wscript.sleep 1000

wshshell.sendkeys "%{F4}"

wscript.sleep 2000

wshshell.sendkeys "{TAB}"

wscript.sleep 1000   

wshshell.sendkeys "{ENTER}"

loop