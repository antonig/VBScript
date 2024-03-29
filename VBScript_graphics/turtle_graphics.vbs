Option Explicit
'SVG image 
'uses a turtle graphics class  to create a sg format file in a html freme and opens it with the default browser
'L-System fractal builder added as example
'pretty coloring of the fractal using hue rotation

'For more L-system fractals: 
'https://fedimser.github.io/l-systems.html
'https://paulbourke.net/fractals/lsys/
function DesktopPath() TempPath=CreateObject("WScript.Shell").SpecialFolders("Desktop") :end function
'to do
' add circle, poligon, polyline
' add filled
' make properties private
'forest: several trees with randomized sizes

' L-System grammar
'  F G	         Move forward by line length drawing a line
'   +	         Turn left by turning angle
'   -	         Turn right by turning angle
'   [	         Push current drawing state onto stack
'   ]	         Pop current drawing state from the stack
'   |	         Reverse direction (ie: turn by 180 degrees)
'   f	         Move forward by line length without drawing a line
' not yat implemented
'   #	         Increment the line width by line width increment
'   !	         Decrement the line width by line width increment
'   @	         Draw a dot with line width radius
'   {	         Open a polygon
'   }	         Close a polygon and fill it with fill colour
'   >	         Multiply the line length by the line length scale factor
'   <	         Divide the line length by the line length scale factor
'   &	         Swap the meaning of + and -
'   (	         Decrement turning angle by turning angle increment
'   )	         Increment turning angle by turning angle increment

Const pi180= 0.01745329251994329576923690768489 ' pi/180 
Const pi=3.1415926535897932384626433832795 'pi
Dim d,turt

function TempPath() 
  TempPath=CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2): 
End function

'
Function hsv2rgb( Hue, Sat, Value) 
  While hue >360. : 
    hue=hue-360:
  Wend
  'hue 0-360   0-ro 120-ver 240-az ,sat 0-100,value 0-100
  Dim Angle, Radius,Ur,Vr,Wr,Rdim
  Dim r,g,b, rgb
  Angle = (Hue-150) *pi180
  Ur = Value * 2.55
  Radius = Ur * tan(Sat *0.01183199)
  Vr = Radius * cos(Angle) *0.70710678  'sqrt(1/2)
  Wr = Radius * sin(Angle) *0.40824829  'sqrt(1/6)
  r = (Ur - Vr - Wr)  
  g = (Ur + Vr - Wr) 
  b = (Ur + Wr + Wr) 
  
  'clamp values 
  If r >255 Then 
    Rdim = (Ur - 255) / (Vr + Wr)
    r = 255
    g = Ur + (Vr - Wr) * Rdim
    b = Ur + 2 * Wr * Rdim 
  ElseIf r < 0 Then
    Rdim = Ur / (Vr + Wr)
    r = 0
    g = Ur + (Vr - Wr) * Rdim
    b = Ur + 2 * Wr * Rdim	
  End If	
  
  If g >255 Then 
    Rdim = (255 - Ur) / (Vr - Wr)
    r = Ur - (Vr + Wr) * Rdim
    g = 255
    b = Ur + 2 * Wr * Rdim
  ElseIf g<0 Then   
    Rdim = -Ur / (Vr - Wr)
    r = Ur - (Vr + Wr) * Rdim
    g = 0
    b = Ur + 2 * Wr * Rdim  	
  End If	
  If b>255 Then
    Rdim = (255 - Ur) / (Wr + Wr)
    r = Ur - (Vr + Wr) * Rdim
    g = Ur + (Vr - Wr) * Rdim
    b = 255
  ElseIf b<0 Then
    Rdim = -Ur / (Wr + Wr)
    r = Ur - (Vr + Wr) * Rdim
    g = Ur + (Vr - Wr) * Rdim
    b = 0
  End If
  hsv2rgb= "#"& right("00"& lcase(hex(r)),2)& right("00"& lcase(hex(g)),2)& right("00"& lcase(hex(b)),2)
End Function

Function countchar(a,x)    'la mas rapida
  countchar = UBound(Split(a,x))
End Function

Class turtle
  Dim fso
  Dim fn
  Dim svg
  
  Dim iang  'radians
  Dim ori   'radians
  Dim incr
  Dim pdown
  Dim clr
  Dim x
  Dim y
  Dim a(16)
  Dim na
  Dim clrinc
  
  ' logo turtle 
  Public Property Let orient(n):
  ori = n*pi180 :
  End Property
  Public Property Let iangle(n):
  iang= n*pi180 :
  End Property
  Public Sub pd() : 
    pdown=True:
  End Sub 
  Public Sub pu()  
    pdown=False :
  End Sub 
  
  Public Sub rt(i)  
    ori=ori - i*iang:
    If ori<0 Then ori = ori+pi*2
  End Sub
  
  Public Sub lt(i):  
    ori=(ori + i*iang) 
    If ori>(pi*2) Then ori=ori-pi*2
  End Sub
  
  
  Public Sub bw(l)
    x= x+ cos(ori+pi)*l*incr
    y= y+ sin(ori+pi)*l*incr
  End Sub 
  
  Public Sub fw(l)
    Dim x1,y1 
    x1=x + cos(ori)*l*incr
    y1=y + sin(ori)*l*incr
    If pdown Then doline x,y,x1,y1
    x=x1:y=y1
  End Sub
  
  
  
  Public Sub push
    a(na)=array(iang,ori,incr,pdown,clr,x,y)
    na=na+1     
  End Sub
  
  Public Sub pop
    If na=0 Then Exit Sub
    na=na-1
    iang=a(na)(0)
    ori=a(na)(1)
    incr=a(na)(2)
    pdown=a(na)(3)
    clr=a(na)(4)
    x=a(na)(5)
    y=a(na)(6)
  End Sub     
  
  ' L-Sys interpreter
  Public Sub dolsys(n,arr)  ' recursion level, definition
    Dim s,s1,i,j,c,hue,inc
    iangle=d("_angle")
    s=d("_start"):s1=""
    
    'build L-sys string
    For i= 1 To n
      For j=1 To len(s)
        c=mid(s,j,1)
        If d.exists(c) Then
          s1=s1 & d(c)
        Else
          s1=s1 & c
        End If
      Next
      s=s1:s1=""
      'wscript.echo i, s  ' displays L-System string. 
    Next
    'draw the fractal
    inc=360/countchar(s,"F") 
    For i=1 To len(s)
      c=mid(s,i,1)
      Select Case c
        Case "+":rt 1
        Case "-":lt 1
        Case "|":ori=ori+pi
        Case "[":push 
        Case "]":pop
        Case "F","G": fw 1 : hue=(hue+inc)  :clr=hsv2rgb(hue,80,70)
        Case "f":pu :fw 1: pd
        Case Else
        ' pending symbols not understood
      End Select
    Next       
  End Sub
  
  
  
  Private Sub Class_Initialize()  
    Set fso = CreateObject ("Scripting.Filesystemobject")
    setlocale "en" 
    'initsvg
    pdown=True
  End Sub
  
  Private Sub Class_Terminate()   
    Set svg=Nothing
    Set fso=Nothing
  End Sub
  
  Public Sub dorect (x,y,x1,y1,c)  'c en hex bgr9
    svg.WriteLine "<rect x=""" & x & """ y= """& y _
    & """ width=""" & x1 & """ height=""" &  y1 & """ style=""fill: "& c & ";"" />"
  End Sub
  
  Public Sub doline (x,y,x1,y1)
    'WScript.Echo "doline"
    svg.WriteLine "<line x1=""" & x & """ y1= """& y & """ x2=""" & x1& """ y2=""" & y1 & """ style=""stroke:"& clr & """/>"
  End Sub 
  
  Public Sub dodot(x,y,r,c)
    svg.WriteLine "<circle cx=""" & x & """ cy= """& y & """ r=""" & r & """  stroke="""& c & """ fill=""" & c &"""/>"
  End Sub
  
  Public Sub display()
    
    svg.writeline "<This browser can't display SVG images. Please update!!>" 
    svg.WriteLine "</svg></body></html>"
    svg.close
    CreateObject("Shell.Application").ShellExecute fn,1,False
  End Sub 
  
  Public Sub initsvg()
    fn=Temppath & "fractal.HTML"
    Set svg = fso.CreateTextFile(fn,True)
    If SVG IS Nothing Then wscript.echo "Can't create svg file" :vscript.quit
    svg.WriteLine "<!DOCTYPE html>" &vbcrlf & "<html>" &vbcrlf & "<head>"
    svg.writeline "<style>" & vbcrlf & "line {stroke:rgb(255,0,0);stroke-width:2}" &vbcrlf &"</style>"
    svg.writeline "</head>"&vbcrlf & "<body>"
    
    svg.WriteLine "<svg xmlns=""http://www.w3.org/2000/svg"" width=""2000"" height=""1000""  style=""background-color:444444""  viewBox=""0 0 1000 1000"">" 
  End Sub 
End Class
Const c_nothing=0
Const c_forward=1



'------------------------------------------------------------------------------------------
Sub dotree
  turt.orient=270
  turt.incr=5
  turt.x=200:turt.y=400
  d.removeall
  d.add "_name","fern"
  d.add "_angle",25
  d.add "_start","X"
  d.add "X","F+[[X]-X]-F[-FX]+X"
  d.add "F","FF"
  turt.dolsys 5,d
End Sub

Sub dosierp
  'WScript.Echo "en sierp"
  'variables : F G
  'constants : + -
  'start  : F-G-G
  'rules  : (F ? F-G+F+G-F), (G ? GG)
  'angle  : 120°
  turt.orient=180
  turt.incr=5
  turt.x=700:turt.y=600	
  d.removeall
  d.add "_name","sierpinski"
  d.add "_angle",120
  d.add "_start","F-G-G"
  d.add "F","F-G+F+G-F"
  d.add "G","GG"
  turt.dolsys 7,d
End Sub

Sub dofern
  Dim x,y,r,n,xn,yn
  For n=1 To 10000
    r = rnd() 'between 0 and 1
    If r < 0.01 Then
      xn = 0.0
      yn = 0.16 * y
    ElseIf r < 0.86 Then
      xn = 0.85 * x + 0.04 * y
      yn = -0.04 * x + 0.85 * y + 1.6
    ElseIf r < 0.93 Then
      xn = 0.2 * x - 0.26 * y
      yn = 0.23 * x + 0.22 * y + 1.6
    Else
      xn = -0.15 * x + 0.28 * y
      yn = 0.26 * x + 0.24 * y + 0.44
    End If   
    turt.dodot xn*100,1000-(yn*90),2,"green"
    x = xn
    y = yn
  Next
End Sub


Sub dohilbert
  ' Alphabet : A, B
  'Constants : F + −
  'Axiom : A
  'Production rules:
  'A → +BF−AFA−FB+
  'B → −AF+BFB+FA− 
  
  turt.orient=0
  turt.incr=4
  turt.x=50:turt.y=50   
  d.RemoveAll
  d.add "_name","hilbert"
  d.add "_angle",90
  d.add "_start","A"
  d.Add  "A","-BF+AFA+FB-"
  d.Add  "B","+AF-BFB-FA+"
  turt.dolsys 7,d
End Sub

Sub dowheel   'wheel is a test of hsv coloring, not a fractal defined in L-SYS
  Dim i
  turt.orient=90
  turt.iangle=5
  turt.incr=5
  turt.pu
  turt.x=500:turt.y=500
  For i=0 To 3590
    turt.dodot turt.x,turt.y,3, hsv2rgb(i mod 360,100-10*i\360,50)
    'turt.dorect turt.x,turt.y,5,5, hsv2rgb(i mod 360,100-10*i\360,50)
    turt.lt 1
    turt.fw 1 +0.001*i
  Next
  turt.pd
End Sub	

Sub dopent   'pentagon
  'axiom = F++F++F++F++F
  'F -> F++F++F|F-F++F
  'angle = 36
  turt.orient=36
  turt.x=30:turt.y=550
  turt.incr=12
  d.removeall
  d.add "_name","pentagon"
  d.add "_angle",36
  d.add "_start","F++F++F++F++F"
  d.add "F","F++F++F|F-F++F"
  turt.dolsys 4,d
End Sub


Sub dodragon
  
  'Axiom: FX
  'Rules:
  'X > X+YF+
  'Y > -FX-Y
  'Angle: 90
  turt.orient=0
  turt.x=300:turt.y=200
  turt.incr=4
  d.removeall
  d.add "_name","dragon"
  d.add "_angle",90
  
  d.add "_start","FX"
  d.Add "X", "X+YF+"
  d.Add "Y" ,"-FX-Y"
  turt.dolsys 15,d
End Sub 


Sub dogosper
  'axiom XF
  'X > X+YF++YF-FX--FXFX-YF+
  'Y > -FX+YFYF++YF+FX--FX-Y
  'angle = 36
  turt.orient=90
  turt.x=100:turt.y=520
  turt.incr=4
  d.removeall
  d.add "_name","gosper"
  d.add "_angle",60
  
  d.add "_start","XF"
  d.add "X","X+YF++YF-FX--FXFX-YF+"
  d.add "Y","-FX+YFYF++YF+FX--FX-Y"
  '  d.Add  "_X",c_nothing
  '  d.Add  "_Y",c_nothing
  '    d.add "_F",c_forward
  turt.dolsys 5,d
End Sub

Sub dosnow
  'axiom = F++F++F
  'F -> F-F++F-F
  'angle = 60
  turt.orient=0
  turt.x=50:turt.y=650
  turt.incr=3
  d.removeall
  d.add "_name","snow"
  d.add "_angle",60
  ' d.add "_F",c_forward
  d.add "_start","F++F++F"
  d.add "F","F-F++F-F"
  turt.dolsys 5,d
End Sub

Sub salir
  WScript.Quit
End Sub  

Function menu(a,t,def)
  'a es array de arrays pares "texto","funcion", en index index 0 cabecera o salir y funcion "salir"
  'pueden ser subs
  Dim i,s,s1,v
  For i=0 To UBound(a)
    s=s&s1&a(i)(0) :s1=vbCrLf&i+1& "- "
  Next
  Do
    Do
      v=InputBox(s,t,def)  'v es texto !!
    Loop until isnumeric(v)
    v=CInt(v) 
  Loop  until  (v >= 0) And (v <= UBound(a))
  menu=a(v)(1)
  'WScript.Echo menu
End Function


' main--------------------------------------

Dim s,x,launch
'the menu
s=Array(array("Select (0=quit):","salir"),Array("wheel","dowheel"),Array(" tree","dotree"), _
Array("pentagon","dopent"),Array("Sierpinski triangle","dosierp"),Array("Hilbert curve","dohilbert") ,_
Array("Gosper curve","dogosper"), Array ("dragon curve","dodragon"),Array("Von Koch Snowflake","dosnow"),_
Array("Beadsley fern","dofern" ))

Set turt=New turtle
Set d=createobject("scripting.dictionary")
Do
  x= menu (s,"Draw Fractals in SVG",0) 'menu pauses the loop so svg is not reset before launching
  turt.initsvg
  Set launch=GetRef(x)
  launch()
  turt.display
Loop 
Set turt=Nothing  'show image in browser
