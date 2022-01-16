Set ws = CreateObject("Wscript.Shell")
Wscript.sleep 800
i=0
do
i=i+1
ws.SendKeys "3"
Wscript.sleep 100
ws.SendKeys "5"
Wscript.sleep 200
ws.SendKeys "{ENTER}"
Wscript.sleep 8000
ws.SendKeys "f"
Wscript.sleep 100
ws.SendKeys "o"
Wscript.sleep 100
ws.SendKeys "r"
Wscript.sleep 100
ws.SendKeys "e"
Wscript.sleep 100
ws.SendKeys "v"
Wscript.sleep 100
ws.SendKeys "e"
Wscript.sleep 100
ws.SendKeys "r"
Wscript.sleep 100
ws.SendKeys ""
Wscript.sleep 100
ws.SendKeys "{Enter}"
Wscript.sleep 8800
loop while i<=9