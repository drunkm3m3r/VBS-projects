Set wshShell =wscript.CreateObject("WScript.Shell")
name=(" type something now")
msgbox(" ") + name
do
wscript.sleep 100
wshshell.sendkeys "{CAPSLOCK}"
wshshell.sendkeys "{NUMLOCK}"
wshshell.sendkeys "{SCROLLLOCK}"
loop