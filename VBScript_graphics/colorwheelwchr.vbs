Option explicit

Class ImgClass
  Private ImgL,ImgH,ImgDepth,bkclr,nclr,tt
  private xmini,xmaxi,ymini,ymaxi
  dim ImgArray()  'rgb in 32 bit mode, indexes to palette in 8 bits
  private filename   
  private Palette
  
public property get xmin():xmin=xmini:end property  
public property get ymin():ymin=ymini:end property  
public property get xmax():xmax=xmaxi:end property  
public property get ymax():ymax=ymaxi:end property  
public property let depth(x)
     if x<>8 and x<>32 then err.raise 9
     Imgdepth=x
end property     
  
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
     depth=dep
     if dep<>8 and dep <>32 then err.raise 9
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

  
  'class termination writes it to a BMP file and displays it 
  'if an error happens VBS terminates the class before exiting so the BMP is displayed the same
  Private Sub Class_Terminate
    
    if err<>0 then wscript.echo "Error " & err.number
    wscript.echo "copying image to bmp file"
    savebmp
    wscript.echo "opening " & filename & " with your default bmp viewer"
    CreateObject("Shell.Application").ShellExecute filename
    wscript.echo timer-tt & " milliseconds"
  End Sub
  
    function long2wstr( x)  'falta muy poco!!!
      dim k1,k2,x1
      k1=  (x and &hffff&)' or (&H8000& And ((X And &h8000&)<>0)))
      k2=((X And &h7fffffff&) \ &h10000&) Or (&H8000& And (x<0))
      long2wstr=chrw(k1) & chrw(k2)
      
    end function 
    
    function int2wstr(x)
        int2wstr=ChrW((x and &h7fff) or (&H8000 And (X<0)))
    End Function


  Public Sub SaveBMP
    'Save the picture to a bmp file
    Dim s,ostream, x,y,loc
    dim bpp:bpp=imgdepth\8
    const hdrs=54 '14+40 
    dim bms:bms=ImgH* 4*(((ImgL*bpp)+3)\4)  'bitmap size including padding
    dim pals:if (imgdepth=8) then pals=(ubound(Palette)+1)*4 else pals=0
    'loc=getlocale
    'setlocale "us"

    with  CreateObject("ADODB.Stream") 'auxiliary ostream, it creates an UNICODE with bom stream in memory
      .Charset = "UTF-16LE"    'o "UTF16-BE" 
      .Type =  2' adTypeText  
      .open 
      
      'build a header
      'bmp header: VBSCript does'nt have records nor writes binary values to files, so we use strings of unicode chars!! 
      'BMP head  0 "BM" 4d42   2 size            6            10                   14
      .writetext ChrW(&h4d42) & long2wstr(hdrs+pals+bms)& long2wstr(0) &long2wstr (hdrs+pals) 

      'InfoHeader 14  hdr sz   18 length      22 width         26 pla       28 clr depth        30 NOCOMPR   34 
      .writetext long2wstr(40) &long2wstr(Imgl)&long2wstr(imgh) & int2wstr(1) & int2wstr(imgdepth)& long2wstr(&H0)
      
       '         34 nosize     38 bpp           42 bpp           46  cls pal     50 imp clrs   54
      .writetext long2wstr(bms)&long2wstr(&Hc4e)& long2wstr(&hc43)& long2wstr(&H0) & long2wstr(&H0)

      'add palette if exists
      If (imgdepth=8) Then
        s=""
        For x=0 to ubound(palette)
          s=s& long2wstr(palette(x))
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
           s=s & long2wstr(Imgarray(x,y))
          Next
          .writetext s
        Next
      Case 8
        dim xx:xx=ImgL mod 4
        For y=ImgH-1 to 0 step-1
          s=""
          For x=0 To ImgL-1 step 2
               s=s & chrw((ImgArray(x,y) and 255 )+ 256*(ImgArray(x+1,y) and 255))
          Next
          if xx and 1 then s=s &chrw(ImgArray(Imgl-1,y))
          if xx >1 then s=s & chrw(0)
          .writetext s 
        Next
      Case Else
        WScript.Echo "ColorDepth not supported : " & ImgDepth & " bits"
      End Select

      'use a second stream to save to file starting past the BOM  the first ADODB.Stream has added
      
      Dim outf:Set outf= CreateObject("ADODB.Stream") 
      outf.Type    = 1 ' adTypeBinary  
      outf.Open
      .position=2              'remove bom (1 wchar) 
      .CopyTo outf
      .close
      outf.savetofile filename,2   'adSaveCreateOverWrite
   
      outf.close
    end with
    'setlocale loc
  End Sub
end class



function hsv2rgb( Hue, Sat, Value) 'hue 0-360   0-ro 120-ver 240-az ,sat 0-100,value 0-100
  dim Angle, Radius,Ur,Vr,Wr,Rdim
  dim r,g,b, rgb
  Angle = (Hue-150) *0.01745329251994329576923690768489
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
 end If
 'b lowest byte, red highest byte
 hsv2rgb= ((b and &hff)+256*((g and &hff)+256*(r and &hff))and &hffffff)
 'hsv2rgb= ((b and &hff)+256*((g and &hff)+256*(r and &hff))and &hffffff)
end function

function ang(col,row)
    'if col =0 then  if row>0 then ang=0 else ang=180:exit function 
    if col =0 then  
      if row<0 then ang=90 else ang=270 end if
    else  
   if col>0 then
      ang=atn(-row/col)*57.2957795130
   else
     ang=(atn(row/-col)*57.2957795130)+180
  end if
  end if
   ang=(ang+360) mod 360  
end function 


Dim X,row,col,fn,tt,hr,sat,row2
const h=160
const w=160
const rad=159
const r2=25500
tt=timer
fn=CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)& "\testwchr.bmp"
Set X = (New ImgClass)(fn,w*2,h*2,32,0,0)

x.set0 w,h
'wscript.echo x.xmax, x.xmin

for row=x.xmin+1 to x.xmax
   row2=row*row
   hr=int(Sqr(r2-row2))
   For col=hr To 159
     Dim a:a=((col\16 +row\16) And 1)* &hffffff
     x.imgArray(col+160,row+160)=a 
     x.imgArray(-col+160,row+160)=a 
   next    
   for col=-hr to hr
     sat=100-sqr(row2+col*col)/rad *50
    ' wscript.echo c,r
     x.imgArray(col+160,row+160)=hsv2rgb(ang(row,col)+90,100,sat)
    next
    'script.echo row
  next  

Set X = Nothing

