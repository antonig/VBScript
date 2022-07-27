option explicit
'outputs turtle graphics to svg file and opens it

'to do
' add circle, poligon, polyline
' add filled
' make properties private

const INVSQR2= 0.70710678
const INVSQR6=  0.40824829

const pi180= 0.01745329251994329576923690768489 ' pi/180 
const pi=3.1415926535897932384626433832795 'pi

'
function hsv2rgb( Hue, Sat, Value) 'hue 0-360   0-ro 120-ver 240-az ,sat 0-100,value 0-100
  dim Angle, Radius,Ur,Vr,Wr,Rdim
  dim r,g,b, rgb
  Angle = (Hue-150) *pi180
  Ur = Value * 2.55
  Radius = Ur * tan(Sat *0.01183199)
  Vr = Radius * cos(Angle) *0.70710678  'sqrt(1/2)
  Wr = Radius * sin(Angle) *0.40824829  'sqrt(1/6)
  r = (Ur - Vr - Wr)  
  g = (Ur + Vr - Wr) 
  b = (Ur + Wr + Wr) 
	
  'clamp values 
 if r >255 then 
   Rdim = (Ur - 255) / (Vr + Wr)
   r = 255
   g = Ur + (Vr - Wr) * Rdim
   b = Ur + 2 * Wr * Rdim 
 elseif r < 0 then
   Rdim = Ur / (Vr + Wr)
   r = 0
   g = Ur + (Vr - Wr) * Rdim
   b = Ur + 2 * Wr * Rdim	
 end if	

 if g >255 then 
   Rdim = (255 - Ur) / (Vr - Wr)
   r = Ur - (Vr + Wr) * Rdim
   g = 255
   b = Ur + 2 * Wr * Rdim
 elseif g<0 then   
   Rdim = -Ur / (Vr - Wr)
   r = Ur - (Vr + Wr) * Rdim
   g = 0
   b = Ur + 2 * Wr * Rdim  	
 end if	
 if b>255 then
   Rdim = (255 - Ur) / (Wr + Wr)
   r = Ur - (Vr + Wr) * Rdim
   g = Ur + (Vr - Wr) * Rdim
   b = 255
 elseif b<0 then
   Rdim = -Ur / (Wr + Wr)
   r = Ur - (Vr + Wr) * Rdim
   g = Ur + (Vr - Wr) * Rdim
   b = 0
 end if
 hsv2rgb= "#"& right("00"& lcase(hex(r)),2)& right("00"& lcase(hex(g)),2)& right("00"& lcase(hex(b)),2)
end function

class turtle
   dim fso
   dim fn
   dim svg
   
   dim iang  'radians
   dim ori   'radians
   dim incr
   dim pdown
   dim clr
   dim x
   dim y
   dim a(16)
   dim na
   public property let orient(n):ori = n*pi180 :end property
   public property let iangle(n):iang= n*pi180 :end property
   public sub pd() : pdown=true: end sub 
   public sub pu()  :pdown=FALSE :end sub 
   
   public sub rt(i)  
     ori=ori - i*iang:
     if ori<0 then ori = ori+pi*2
   end sub 
   public sub lt(i):  
     ori=(ori + i*iang) 
     if ori>(pi*2) then ori=ori-pi*2
   end sub
   
   public sub push
     a(na)=array(iang,ori,incr,pdown,clr,x,y)
     na=na+1     
   end sub
   public sub pop
     if na=0 then exit sub
     na=na-1
     iang=a(na)(0)
     ori=a(na)(1)
     incr=a(na)(2)
     pdown=a(na)(3)
     clr=a(na)(4)
     x=a(na)(5)
     y=a(na)(6)
   end sub     
   
   public sub dolsys(n,arr)
   dim s,s1,i,j,c 
	     iangle=d("_angle")
       s=d("_start"):s1=""
       'rewrite
       for i= 1 to n
         for j=1 to len(s)
            c=mid(s,j,1)
            if d.exists(c) then
              s1=s1 & d(c)
            else
              s1=s1 & c
            end if
          next
          s=s1:s1=""
          wscript.echo i, s
       next
       'draw
			 
       for i=1 to len(s)
         c=mid(s,i,1)
         select case c
         case "+":rt 1
         case "-":lt 1
				 case "|":ori=ori+pi
				 case "[":push 
         case "]":pop
         case else
           dim dc:dc ="_"& c
           if d.exists(dc) then
             select case  d(dc)
             case c_nothing:
             case c_forward:fw 1
             case else
                 wscript.echo "command " & c & " not listed"
             end select  
           else
              wscript.echo "variable " & c & " not listed"
           end if
         end select
      next       
   end sub
   
   
   public sub bw(l)
      x= x+ cos(ori+pi)*l*incr
      y= y+ sin(ori+pi)*l*incr
   end sub 
   
   public sub fw(l)
      dim x1,y1 
      x1=x + cos(ori)*l*incr
      y1=y + sin(ori)*l*incr
      if pdown then doline x,y,x1,y1
      x=x1:y=y1
   end sub
   
   Private Sub Class_Initialize()  
     setlocale "en" 
      initsvg
      pdown=true
   end sub
   
   Private Sub Class_Terminate()   
      disply
   end sub
   
   public sub dorect (x,y,x1,y1,c)  'c en hex bgr9
      svg.WriteLine "<rect x=""" & x & """ y= """& y _
      & """ width=""" & x1 & """ height=""" &  y1 & """ style=""fill: "& c & ";"" />"
   end sub
   
   private sub doline (x,y,x1,y1)
      svg.WriteLine "<line x1=""" & x & """ y1= """& y & """ x2=""" & x1& """ y2=""" & y1 & """/>"
   end sub 

   private sub disply()
       dim shell
			 svg.writeline "This browswr can't display SVG images. Please update!!" 
       svg.WriteLine "</svg></body></html>"
       svg.close
       Set shell = CreateObject("Shell.Application") 
       shell.ShellExecute fn,1,False
   end sub 

   private sub initsvg()
     dim scriptpath
     Set fso = CreateObject ("Scripting.Filesystemobject")
     ScriptPath= Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
     fn=Scriptpath & "SIERP.HTML"
     Set svg = fso.CreateTextFile(fn,True)
     if SVG IS nothing then wscript.echo "Can't create svg file" :vscript.quit
     svg.WriteLine "<!DOCTYPE html>" &vbcrlf & "<html>" &vbcrlf & "<head>"
     svg.writeline "<style>" & vbcrlf & "line {stroke:rgb(255,0,0);stroke-width:.5}" &vbcrlf &"</style>"
     svg.writeline "</head>"&vbcrlf & "<body>"
     svg.WriteLine "<svg xmlns=""http://www.w3.org/2000/svg"" width=""1000"" height=""1000"" viewBox=""0 0 1000 1000"">" 
   end sub 
end class
const c_nothing=0
const c_forward=1

sub dofern
  x.orient=270
  x.incr=5
  x.x=100:x.y=300
	d.removeall
  d.add "_name","fern"
  d.add "_angle",25
  d.add "_X",c_nothing
  d.add "_F",c_forward
  d.add "_start","X"
  d.add "X","F+[[X]-X]-F[-FX]+X"
  d.add "F","FF"
  x.dolsys 3,d
end sub

sub dosierp
'variables : F G
'constants : + -
'start  : F-G-G
'rules  : (F ? F-G+F+G-F), (G ? GG)
'angle  : 120°
x.x=200:x.y=100
  x.incr=6
	
 d.removeall
 d.add "_name","sierpinski"
 d.add "_angle",120
 d.add "_F",c_forward
 d.add "_G",c_forward
 d.add "_start","F-G-G"
 d.add "F","F-G+F+G-F"
 d.add "G","GG"
 x.dolsys 7,d
end sub


sub dowheel
  x.orient=90
	x.iangle=5
	x.incr=5
	x.pu
	x.x=500:x.y=500
	for i=0 to 3590
	  x.dorect x.x,x.y,5,5, hsv2rgb(i mod 360,100-10*i\360,50)
	  x.lt 1
	  x.fw 1+0.001*i
	next
	x.pd
end sub	

sub dopent
'axiom = F++F++F++F++F
'F -> F++F++F|F-F++F
'angle = 36
  x.orient=90
  x.x=500:x.y=500
  x.incr=16
 d.removeall
 d.add "_name","pentagon"
 d.add "_angle",36
 d.add "_F",c_forward
  d.add "_start","F++F++F++F++F"
 d.add "F","F++F++F|F-F++F"
 x.dolsys 4,d
end sub

 dim x,i,d
 set d=createobject("scripting.dictionary")
set x=new turtle
'dowheel
'dofern 
'dopent
dosierp
set x=nothing  'show image in browser


