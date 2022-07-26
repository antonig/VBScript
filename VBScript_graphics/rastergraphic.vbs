option explicit
' pdte
'   ejes negativos, 
'   limitar bmp maximo
'ok pasar a 32 bits chrw, buffer de linea y stream para salvar mas rapido
'   clipping de circunferencia, arco elipse, quesito
'   polilineas, poligonos
'   turtle
'   linea desde
'   floodfill

' Raster graphics class in VBSCRIPT by Antoni Gual
'--------------------------------------------
' An array keeps the image allowing to set pixels, draw lines and boxes in it. 
' at class destroy a bmp file is saved to disk and the default viewer is called
' The class can work with 8 and 24 bit bmp. With 8 bit uses a built-in palette or can import user's


'Declaration : 
' Set MyObj = (New ImgClass)(name,width,hight,bits_per_pixel,background_clr,palette_array)
' name:path and name of the file
' bits_per_pixel can bs only 8 and 24
' palette array only to substitute the default palette for 8 bits, else put a 0


'Properties :
'    Pixel(x,y) R/W, x=0..ImgL-1, y=0..ImgH-1. Clipping Get/set the color-code of a pixel.
'    if clipping not needed the array ImgArray can be read/written directly
'Methods:
     
'    Line (x0,y0,x1,y1,c)  draws a line (x0,y0) to (x1,y1) in the color c 
'    boxf (x0,y0,w,h,c)    draws abox wit top left corner at x0,y0 filled with color c
'    circle (x0,y0,rad,c)  draws a circle with center (x0,y0) radius r color c
'    c must be a byte index to the palette in 8 bit images and a rgb value for 24 bit images

'
'to do:
'    fat line 
'    circle clipping
'    orientacion
'    draw axles
'    graph points
'    gradients
'    load bmp
'    allow loading of palettes <256
'    background image loading
'    dithering
'    point (cross, circle,triangle,square)
'    ellipse, arc, cheese
'    count colors in 24 bit mode 
'    RLE for saving 8 bit images


const cl_INSIDE = 0 
const cl_LEFT = 1   
const cl_RIGHT = 2  ' 0010
const  cl_BOTTOM = 4 ' 0100
const  cl_TOP = 8    ' 1000


Class ImgClass
  Private ImgL,ImgH,ImgDepth,bkclr,loc,tt
  private xmini,xmaxi,ymini,ymaxi
  private ImgArray()  'rgb in 24 bit mode, indexes to palette in 8 bits
  private filename   
  private Palette
  
public property get xmin():xmin=xmini:end property  
public property get ymin():ymin=ymini:end property  
public property get xmax():xmax=xmaxi:end property  
public property get ymax():ymax=ymaxi:end property  

  Public Property Let Pixel (x,y,color)  'clipping
    If (x=<xmaxi) And (x>=xmini) And (y<=ymaxi) And (y>=xmini) Then
        ImgArray(x-xmini,y-ymini)=Color
    end if             
  End Property

  Public Property Get Pixel (x,y)  'clipping
    If (x=<xmaxi) And (x>=xmini) And (y<=ymaxi) And (y>=xmini) Then
      Pixel=ImgArray(x-xmini,y-ymaxi)
    else
      pixel=0   
    End If
  End Property
  
  
  public sub set0 (x0,y0) 'sets the new origin (default tlc)
    if x0<0 or x0>=imgl or y0<0 or y0>imgh then err.raise 9 
    xmini=-x0
    ymini=-y0
    xmaxi=xmini+imgl-1
    ymaxi=ymini+imgh-1    
  end sub
  
  'constructor
  Public Default Function Init(name,w,h,dep,bkg,pal)
     dim i,j
     ImgL=w
     ImgH=h
     tt=timer
      loc=getlocale
     set0 0,0  'tlc
     redim imgArray(ImgL-1,ImgH-1)
     bkclr=bkg
     if bkg<>0 then 
       for i=0 to ImgL-1 
         for j=0 to ImgH-1 
            imgarray(i,j)=bkg
         next
       next  
    end if   
     filename=name
     ImgDepth =dep
     'load user palette if provided  
     if imgdepth=8 then 
       if isarray(pal) then
        if ubound(pal)=255 then
            palette=pal
        else
           mypalette          
        end if
      else
        mypalette
      end if 
     end if       
     set init=me
  end function

  private sub mypalette
    'Default palette recycled from ATARI
    palette=array(_
    &h000000, &h111111, &h222222, &h333333, &h444444, &h555555, &h666666, &h777777, &h888888, &h999999, &hAAAAAA, &hBBBBBB, &hCCCCCC, &hDDDDDD, &hEEEEEE, &hFFFFFF,_
    &h003200, &h114300, &h225400, &h336500, &h447600, &h558700, &h669800, &h77A900, &h88BA00, &h99CB00, &hAADC00, &hBBED00, &hCCFE00, &hDDFF00, &hEEFF00, &hFFFF00,_
    &h3B1000, &h4C2100, &h5D3200, &h6E4300, &h7F5400, &h906500, &hA17600, &hB28700, &hC39800, &hD4A900, &hE5BA00, &hF6CB00, &hFFDC00, &hFFED00, &hFFFE01, &hFFFF12,_
    &h6C0000, &h7D0000, &h8E0D00, &h9F1E00, &hB02F00, &hC14000, &hD25100, &hE36200, &hF47300, &hFF8400, &hFF9500, &hFFA60E, &hFFB71F, &hFFC830, &hFFD941, &hFFEA52,_
    &h8A0000, &h9B0000, &hAC0000, &hBD0000, &hCE0D00, &hDF1E05, &hF02F16, &hFF4027, &hFF5138, &hFF6249, &hFF735A, &hFF846B, &hFF957C, &hFFA68D, &hFFB79E, &hFFC8AF,_
    &h91001B, &hA2002C, &hB3003D, &hC4004E, &hD5005F, &hE60670, &hF71781, &hFF2892, &hFF39A3, &hFF4AB4, &hFF5BC5, &hFF6CD6, &hFF7DE7, &hFF8EF8, &hFF9FFF, &hFFB0FF,_
    &h7E0082, &h8F0093, &hA000A4, &hB100B5, &hC200C6, &hD300D7, &hE40DE8, &hF51EF9, &hFF2FFF, &hFF40FF, &hFF51FF, &hFF62FF, &hFF73FF, &hFF84FF, &hFF95FF, &hFFA6FF,_
    &h5500D2, &h6600E3, &h7700F4, &h8800FF, &h9900FF, &hAA01FF, &hBB12FF, &hCC23FF, &hDD34FF, &hEE45FF, &hFF56FF, &hFF67FF, &hFF78FF, &hFF89FF, &hFF9AFF, &hFFABFF,_
    &h1E00FD, &h2F00FF, &h4000FF, &h5100FF, &h6203FF, &h7314FF, &h8425FF, &h9536FF, &hA647FF, &hB758FF, &hC869FF, &hD97AFF, &hEA8BFF, &hFB9CFF, &hFFADFF, &hFFBEFF,_
    &h0000FD, &h0000FF, &h0400FF, &h1511FF, &h2622FF, &h3733FF, &h4844FF, &h5955FF, &h6A66FF, &h7B77FF, &h8C88FF, &h9D99FF, &hAEAAFF, &hBFBBFF, &hD0CCFF, &hE1DDFF,_
    &h0003D2, &h0014E3, &h0025F4, &h0036FF, &h0047FF, &h0058FF, &h1169FF, &h227AFF, &h338BFF, &h449CFF, &h55ADFF, &h66BEFF, &h77CFFF, &h88E0FF, &h99F1FF, &hAAFFFF,_
    &h002782, &h003893, &h0049A4, &h005AB5, &h006BC6, &h007CD7, &h008DE8, &h009EF9, &h0AAFFF, &h1BC0FF, &h2CD1FF, &h3DE2FF, &h4EF3FF, &h5FFFFF, &h70FFFF, &h81FFFF,_
    &h00441B, &h00552C, &h00663D, &h00774E, &h00885F, &h009970, &h00AA81, &h00BB92, &h00CCA3, &h08DDB4, &h19EEC5, &h2AFFD6, &h3BFFE7, &h4CFFF8, &h5DFFFF, &h6EFFFF,_
    &h005600, &h006700, &h007800, &h008900, &h009A00, &h00AB05, &h00BC16, &h00CD27, &h00DE38, &h0FEF49, &h20FF5A, &h31FF6B, &h42FF7C, &h53FF8D, &h64FF9E, &h75FFAF,_
    &h005900, &h006A00, &h007B00, &h008C00, &h009D00, &h00AE00, &h00BF00, &h0BD000, &h1CE100, &h2DF200, &h3EFF00, &h4FFF0E, &h60FF1F, &h71FF30, &h82FF41, &h93FF52,_
    &h004C00, &h005D00, &h006E00, &h007F00, &h099000, &h1AA100, &h2BB200, &h3CC300, &h4DD400, &h5EE500, &h6FF600, &h80FF00, &h91FF00, &hA2FF00, &hB3FF01, &hC4FF12_
     )
  End Sub
  
  'class termination writes it to a BMP file and displays it 
  'if an error happens VBS terminates the class before exiting so the BMP is displayed the same
  Private Sub Class_Terminate
    savebmp
    wscript.echo "opening " & filename
    CreateObject("Shell.Application").ShellExecute filename
    wscript.echo "Tiempo " & timer-tt&"ms"
  End Sub

  'writes a 32bit integr value as binary to a string
  function long2str(byval k)
        Dim s
        s= chr(k and &hff)
        k=k\&h100
        s=s& chr(k and &hff)
        k=k\&h100
        s=s& chr(k and &hff)
        k=k\&h100
        s=s& chr(k and &hff)
        long2str=s
    End function
  
    function int2str(byval k)
        Dim s
        s= chr(k and &hff)
        k=k\&h100
        s=s& chr(k and &hff)
        int2str=s
    End function
  
  

  Public Sub SaveBMP
    'Save the picture to a bmp file
    Dim s,ostream, x,y
    dim bpp:bpp=imgdepth\8
    const hdrs=54 '14+40 
    dim bms:bms=ImgH* 4*(((ImgL*bpp)+3)\4)  'bitmap size including padding
    dim pals:if (imgdepth=8) then pals=(ubound(Palette)+1)*4 else pals=0
    setlocale "en-us" 

    with  CreateObject("ADODB.Stream") 'auxiliary ostream, it creates an UNICODE with bomb stream in memory
      .Charset = "Windows-1252"  
      .Type =  2' adTypeText  
      .open 
      
      'build a header
      'bmp header: VBSCript does'nt have records nor writes binary values to files, so we use strings of unicode chars!! 
      'BMP head  0 "BM" 'Type  2 size            6            10                   14
      .writetext "BM" & long2str(hdrs+pals+bms)& long2str(0) &long2str (hdrs+pals) 

      'InfoHeader 14  hdr sz   18 length      22 width         26 pla       28 clr depth        30 NOCOMPR   34 
      .writetext long2str(40) &long2str(Imgl)&long2str(imgh) & int2str(1) & int2str(imgdepth)& long2str(&H0)
      
       '         34 nosize     38 bpp           42 bpp           46  cls pal     50 imp clrs   54
      .writetext long2str(bms)&long2str(&Hc4e)& long2str(&hc43)& long2str(&H0) & long2str(&H0)

      'add palette if exists
      If (imgdepth=8) Then
        s=""
        For x=0 to ubound(palette)
          s=s& long2str(palette(x))
        Next
        .writetext s
      End If
      
      'write bitmap
      Select Case ImgDepth
      Case 32
      'wscript.echo imgdepth
        For y=ImgH-1 to 0 step-1  'Origin of bitmap: bottom left
          s=""
          For x=0 To ImgL-1
           'writelong fic, Pixel(x,y) 
           s=s & long2str(Imgarray(x,y))
          Next
          .writetext s
        Next
      Case 8
        dim xx:xx=ImgL mod 4
        dim align:align=string(xx,0)
        For y=ImgH-1 to 0 step-1
          s=""
          For x=0 To ImgL-1 
               s=s & chr(ImgArray(x,y) and 255 )
          Next
          if xx <>0 then s=s & align
          .writetext s 
        Next
  
      Case Else
        WScript.Echo "ColorDepth not supported : " & ImgDepth & " bits"
      End Select

      'save to file
      .position=0              'using single byte chars, no bom
      .savetofile filename,2   'adSaveCreateOverWrite
      .close
    end with
     setlocale loc
  End Sub

 
 'bresenham's line, does not clip 
 private Sub linenc(byval x0,byval y0, byval x1, byval y1,c)  
    Dim x,y,xf,yf,dx,dy,sx,sy,err,err2
    x =cint(x0-xmini)     : y =cint(y0-ymini)
    xf=cint(x1-ymini)     : yf=cint(y1-ymini)
    dx=Abs(xf-x)          : dy=Abs(yf-y)
   
    'wscript.echo ">",x,y,xf,yf,dx,dy
    If x<xf Then sx=+1: Else sx=-1
    If y<yf Then sy=+1: Else sy=-1
    err=dx-dy
    Do
      'wscript.echo x,y,err,sx,sy
      ImgArray(x,y)=c    
      If x=xf And y=yf Then Exit Do
      err2=err+err
      If err2>-dy Then err=err-dy: x=x+sx
      If err2< dx Then err=err+dx: y=y+sy
    Loop

  End Sub 'draw_line 

'Cohen-sutherland  line clipping
private function ComputeCode(x, y)
  dim code: code = cl_INSIDE   ' initialised as being inside of window
  if (x < xmini) then          ' to the left of clip window
    code =code or cl_LEFT
  elseif (x > xmaxi) then      ' to the right of clip window
    code = code or cl_RIGHT
  end if  
  if (y < ymini) then          ' below the clip window
    code = code or cl_BOTTOM
  elseif (y > ymaxi) then      ' above the clip window
    code = code or cl_TOP
  end if
  computeCode= code
end function

'line drawing using Cohen-sutherland  clipping
  public sub line(byval x0, byval y0, byval x1, byval y1, clr) 'clipping
    dim outcode0: outcode0 = ComputeCode(x0, y0)
    dim outcode1: outcode1 = ComputeCode(x1, y1)
    dim ok   '
    do while 1 'loop exit ok if both ends clip as visible, not ok if line not visible at all 
    'wscript.echo outcode0,outcode1
    if (outcode0 or outcode1)=0 then  'both ends visible 
        ok=true :exit do
    elseif (outcode0 and outcode1) then  'both ends at one side,line not visible
        ok=false:exit do
    else 
      dim  x, y,outcodeout,slope
      if outcode1  then outcodeout= outcode1 else outcodeout= outcode0
      if (outcodeOut and  cl_TOP)then           
        x = int(x0 + (x1 - x0) * (ymaxi - y0) / (y1 - y0))
        y = ymaxi
      elseif (outcodeOut and cl_BOTTOM) then  ' point is below the clip window
        x = int(x0 + (x1 - x0) * (ymini - y0) / (y1 - y0))
        y = ymini
      elseif (outcodeOut and cl_RIGHT) then   ' point is to the right of clip window
        y = int(y0 + (y1 - y0) * (xmaxi - x0) / (x1 - x0))
        x = xmaxi
      elseif (outcodeOut and cl_LEFT) then   ' point is to the left of clip window
        y = int(y0 + (y1 - y0) * (xmini - x0) / (x1 - x0))
        x = xmini
      end if
      if (outcodeOut= outcode0) then 
        x0 = x:y0 = y  
        outcode0 = ComputeCode(x0, y0)
      else 
        x1 = x:y1 = y
        outcode1 = ComputeCode(x1, y1)
      end if  
    end if
    loop
     'wscript.echo x0,y0,x1,y1,clr, ok
    if ok then linenc x0, y0, x1, y1 ,clr  
  end sub  

  public sub fbox (x0,y0, w,h,byval c) 'filled box,clips  
    dim i,j 
    dim x1,y1,x2,y2
    'clipping
    if x0<xmini then x1=xmini else x1=x0 
    if y0<ymini then y1=xmaxi else y1=y0 
    if x0+w>=xmaxi then x2=ymaxi else x2=x0+w 
    if y0+h>=ymaxi then y2=ymaxi else y2=y0+h 
      for i = x1-xmini to x2-xmini  
        for j = y1-xmini to y2-xmini 
          ImgArray(i,j)=c
        next
      next
  end sub
  
  public sub circle(x1,y1,r,c) 'clips pixel by pixel
    dim x,y,err,x0,y0 
    x=r:y=0:err=0:
    x0=x1:y0=y1      'no sumar xmini ymini porque pixel ya lo hace
    wscript.echo x0,y0
    do while x>=y 
      pixel (x0 + x, y0 + y)=c
      pixel (x0 + y, y0 + x)=c
      pixel (x0 - y, y0 + x)=c
      pixel (x0 - x, y0 + y)=c
      pixel (x0 - x, y0 - y)=c
      pixel (x0 - y, y0 - x)=c
      pixel (x0 + y, y0 - x)=c
      pixel (x0 + x, y0 - y)=c

      y = y+ 1
      if err <= 0 then err =err+ 2*y + 1
      if err > 0 then x =x- 1:err = err- 2*x + 1
    loop  
  end sub
End Class
const pi180= 0.01745329251994329576923690768489 ' pi/180 
const pi=3.1415926535897932384626433832795 'pi

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
 hsv2rgb= (b and &hff)+256*((g and &hff)+256*(r and &hff))

end function


sub test32
'squiggly
Dim X,i,j,fn,t
const h=240
const w=320

fn=CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)& "\test32.bmp"
Set X = (New ImgClass)(fn,w*2,h*2,32,0,0)
x.set0 w,h
const scal=120
For t = 0 To 4*Pi Step .01
  x.Line scal*cos(t)+30, -scal*sin(t),30+scal*(cos(t) - sin(2*t)), -scal*(sin(t) + cos(t/2)), hsv2rgb(t*28.647889,90,50)
Next
x.pixel (0,0)=&hffffff
Set X = Nothing
end sub


sub test8
'palette
Dim X,i,j,fn,t
const sq=20 
dim h: h=16*sq
dim w: w=16*sq
fn=CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)& "\test8.bmp"
'display palette

Set X = (New ImgClass)(fn,w,h,8,0,0)

'palette
for i= 0 to 15
  for j=0 to 15
    x.fbox i*sq-x.xmin,j*sq-x.ymin,sq,sq,(j*16+i)
  next
next

x.set0 w\2,h\2
x.circle 0,0,150,16
x.circle 0,0,140,0
x.line 22,0,400,x.ymax,&h47
x.line -7,-40,500,347,&h32

Set X = Nothing
end sub

test32
test8