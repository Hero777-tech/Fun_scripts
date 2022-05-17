Set wshShell =wscript.CreateObject("WScript.Shell") 'disco dewane - created by Adtiya-2019
name=(" Watch the Magic on your Keyboard ")
msgbox(" ") + name          'keyboard disco
do
wscript.sleep 100
wshshell.sendkeys "{CAPSLOCK}"
wshshell.sendkeys "{NUMLOCK}"
wshshell.sendkeys "{SCROLLLOCK}"
loop