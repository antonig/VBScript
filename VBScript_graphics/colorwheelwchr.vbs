Option explicit

Class ImgClass
  Private ImgL,ImgH,ImgDepth,bkclr,loc,tt
  private xmini,xmaxi,ymini,ymaxi,dirx,diry
  public ImgArray()  'rgb in 24 bit mode, indexes to palette in 8 bits
  private filename   
  private Palette,szpal 
  
  public property get xmin():xmin=xmini:end property  
  public property get ymin():ymin=ymini:end property  
  public property get xmax():xmax=xmaxi:end property  
  public property get ymax():ymax=ymaxi:end property  
  public property let depth(x)
  if x<>8 and x<>32 then err.raise 9
  Imgdepth=x
  end property     
  
  public sub set0 (x0,y0) 'sets the new origin (default tlc). The origin does'nt work if ImgArray is accessed directly
    if x0<0 or x0>=imgl or y0<0 or y0>imgh then err.raise 9 
    xmini=-x0
    ymini=-y0
    xmaxi=xmini+imgl-1
    ymaxi=ymini+imgh-1    
  end sub
  
  'constructor
  Public Default Function Init(name,w,h,orient,dep,bkg,mipal)
  'offx, offy posicion de 0,0. si ofx+ , x se incrementa de izq a der, si offy+ y se incrementa de abajo arriba
  dim i,j
  ImgL=w
  ImgH=h
  tt=timer
  loc=getlocale
  ' not useful as we are not using SetPixel and accessing  ImgArray directly
  set0 0,0   'origin blc positive up and right
  redim imgArray(ImgL-1,ImgH-1)
  bkclr=bkg
  if bkg<>0 then 
    for i=0 to ImgL-1 
      for j=0 to ImgH-1 
        imgarray(i,j)=bkg
      next
    next  
  end if 
  Select Case orient
    Case 1: dirx=1 : diry=1   
    Case 2: dirx=-1 : diry=1
    Case 3: dirx=-1 : diry=-1
    Case 4: dirx=1 : diry=-1
  End select    
  filename=name
  ImgDepth =dep
  'load user palette if provided  
  if imgdepth=8 then  
    loadpal(mipal)
  end if       
  set init=me
  end function

  private sub loadpal(mipale)
    if isarray(mipale) Then
      palette=mipale
      szpal=UBound(mipale)+1
    Else
      szpal=256  
    'Default palette recycled from ATARI
    palette=array(_
    &h000000&, &h111111&, &h222222&, &h333333&, &h444444&, &h555555&, &h666666&, &h777777&, &h888888&, &h999999&, &hAAAAAA&, &hBBBBBB&, &hCCCCCC&, &hDDDDDD&, &hEEEEEE&, &hFFFFFF&,_
    &h003200&, &h114300&, &h225400&, &h336500&, &h447600&, &h558700&, &h669800&, &h77A900&, &h88BA00&, &h99CB00&, &hAADC00&, &hBBED00&, &hCCFE00&, &hDDFF00&, &hEEFF00&, &hFFFF00&,_
    &h3B1000&, &h4C2100&, &h5D3200&, &h6E4300&, &h7F5400&, &h906500&, &hA17600&, &hB28700&, &hC39800&, &hD4A900&, &hE5BA00&, &hF6CB00&, &hFFDC00&, &hFFED00&, &hFFFE01&, &hFFFF12&,_
    &h6C0000&, &h7D0000&, &h8E0D00&, &h9F1E00&, &hB02F00&, &hC14000&, &hD25100&, &hE36200&, &hF47300&, &hFF8400&, &hFF9500&, &hFFA60E&, &hFFB71F&, &hFFC830&, &hFFD941&, &hFFEA52&,_
    &h8A0000&, &h9B0000&, &hAC0000&, &hBD0000&, &hCE0D00&, &hDF1E05&, &hF02F16&, &hFF4027&, &hFF5138&, &hFF6249&, &hFF735A&, &hFF846B&, &hFF957C&, &hFFA68D&, &hFFB79E&, &hFFC8AF&,_
    &h91001B&, &hA2002C&, &hB3003D&, &hC4004E&, &hD5005F&, &hE60670&, &hF71781&, &hFF2892&, &hFF39A3&, &hFF4AB4&, &hFF5BC5&, &hFF6CD6&, &hFF7DE7&, &hFF8EF8&, &hFF9FFF&, &hFFB0FF&,_
    &h7E0082&, &h8F0093&, &hA000A4&, &hB100B5&, &hC200C6&, &hD300D7&, &hE40DE8&, &hF51EF9&, &hFF2FFF&, &hFF40FF&, &hFF51FF&, &hFF62FF&, &hFF73FF&, &hFF84FF&, &hFF95FF&, &hFFA6FF&,_
    &h5500D2&, &h6600E3&, &h7700F4&, &h8800FF&, &h9900FF&, &hAA01FF&, &hBB12FF&, &hCC23FF&, &hDD34FF&, &hEE45FF&, &hFF56FF&, &hFF67FF&, &hFF78FF&, &hFF89FF&, &hFF9AFF&, &hFFABFF&,_
    &h1E00FD&, &h2F00FF&, &h4000FF&, &h5100FF&, &h6203FF&, &h7314FF&, &h8425FF&, &h9536FF&, &hA647FF&, &hB758FF&, &hC869FF&, &hD97AFF&, &hEA8BFF&, &hFB9CFF&, &hFFADFF&, &hFFBEFF&,_
    &h0000FD&, &h0000FF&, &h0400FF&, &h1511FF&, &h2622FF&, &h3733FF&, &h4844FF&, &h5955FF&, &h6A66FF&, &h7B77FF&, &h8C88FF&, &h9D99FF&, &hAEAAFF&, &hBFBBFF&, &hD0CCFF&, &hE1DDFF&,_
    &h0003D2&, &h0014E3&, &h0025F4&, &h0036FF&, &h0047FF&, &h0058FF&, &h1169FF&, &h227AFF&, &h338BFF&, &h449CFF&, &h55ADFF&, &h66BEFF&, &h77CFFF&, &h88E0FF&, &h99F1FF&, &hAAFFFF&,_
    &h002782&, &h003893&, &h0049A4&, &h005AB5&, &h006BC6&, &h007CD7&, &h008DE8&, &h009EF9&, &h0AAFFF&, &h1BC0FF&, &h2CD1FF&, &h3DE2FF&, &h4EF3FF&, &h5FFFFF&, &h70FFFF&, &h81FFFF&,_
    &h00441B&, &h00552C&, &h00663D&, &h00774E&, &h00885F&, &h009970&, &h00AA81&, &h00BB92&, &h00CCA3&, &h08DDB4&, &h19EEC5&, &h2AFFD6&, &h3BFFE7&, &h4CFFF8&, &h5DFFFF&, &h6EFFFF&,_
    &h005600&, &h006700&, &h007800&, &h008900&, &h009A00&, &h00AB05&, &h00BC16&, &h00CD27&, &h00DE38&, &h0FEF49&, &h20FF5A&, &h31FF6B&, &h42FF7C&, &h53FF8D&, &h64FF9E&, &h75FFAF&,_
    &h005900&, &h006A00&, &h007B00&, &h008C00&, &h009D00&, &h00AE00&, &h00BF00&, &h0BD000&, &h1CE100&, &h2DF200&, &h3EFF00&, &h4FFF0E&, &h60FF1F&, &h71FF30&, &h82FF41&, &h93FF52&,_
    &h004C00&, &h005D00&, &h006E00&, &h007F00&, &h099000&, &h1AA100&, &h2BB200&, &h3CC300&, &h4DD400&, &h5EE500&, &h6FF600&, &h80FF00&, &h91FF00&, &hA2FF00&, &hB3FF01&, &hC4FF12&_
     )
   End if  
  End Sub

  
  'class termination writes it to a BMP file and displays it 
  'if an error happens VBS terminates the class before exiting so the BMP is displayed the same
  Private Sub Class_Terminate
    
    if err<>0 then wscript.echo "Error " & err.number
    wscript.echo "copying image to bmp file"
    savebmp
    wscript.echo "opening " & filename & " with your default bmp viewer"
    CreateObject("Shell.Application").ShellExecute filename
    wscript.echo timer-tt & "  iseconds"
  End Sub
  
 'writes a 32bit integr value as binary to an utf16 string

function long2wstr( x) 
   long2wstr=chrw(x and &hffff&) + ChrW(((X And &h7fffffff&) \ &h10000&) Or (&H8000& And (x<0))) 
end Function 
    
    function int2wstr(x)
        int2wstr=ChrW((x and &h7fff) or (&H8000& And (X<0)))
    End Function


   Public Sub SaveBMP
    'Save the picture to a bmp file
    Dim s,ostream, x,y,loc
   
    const hdrs=54 '14+40 
    dim bms:bms=ImgH* 4*(((ImgL*imgdepth\8)+3)\4)  'bitmap size including padding
    dim palsize:if (imgdepth=8) then palsize=szpal*4 else palsize=0

    with  CreateObject("ADODB.Stream") 'auxiliary ostream, it creates an UNICODE with bom stream in memory
      .Charset = "UTF-16LE"    'o "UTF16-BE" 
      .Type =  2' adTypeText  
      .open 
      
      'build a header
      'bmp header: VBSCript does'nt have records nor writes binary values to files, so we use strings of unicode chars!! 
      'BMP header  
      .writetext ChrW(&h4d42)                           ' 0 "BM" 4d42 
      .writetext long2wstr(hdrs+palsize+bms)            ' 2 fiesize  
      .writetext long2wstr(0)                           ' 6  reserved 
      .writetext long2wstr (hdrs+palsize)               '10 image offset 
       'InfoHeader 
      .writetext long2wstr(40)                          '14 infoheader size
      .writetext long2wstr(Imgl)                        '18 image length  
      .writetext long2wstr(imgh)                        '22 image width
      .writetext int2wstr(1)                            '26 planes
      .writetext int2wstr(imgdepth)                     '28 clr depth (bpp)
      .writetext long2wstr(&H0)                         '30 compression used 0= NOCOMPR
       
      .writetext long2wstr(bms)                         '34 imgsize
      .writetext long2wstr(&Hc4e)                       '38 bpp hor
      .writetext long2wstr(&hc43)                       '42 bpp vert
      .writetext long2wstr(szpal)                       '46  colors in palette
      .writetext long2wstr(&H0)                         '50 important clrs 0=all
     
      'write bitmap
      'precalc data for orientation
       Dim x1,x2,y1,y2
       If dirx=-1 Then x1=ImgL-1 :x2=0 Else x1=0:x2=ImgL-1
       If diry=-1 Then y1=ImgH-1 :y2=0 Else y1=0:y2=ImgH-1 
       
      Select Case imgdepth
      
      Case 32
        For y=y1 To y2  step diry   
          For x=x1 To x2 Step dirx
           'writelong fic, Pixel(x,y) 
           .writetext long2wstr(Imgarray(x,y))
          Next
        Next
        
      Case 8
        'palette
        For x=0 to szpal-1
          .writetext long2wstr(palette(x))  '52
        Next
        'image
        dim pad:pad=ImgL mod 4
        For y=y1 to y2 step diry
          For x=x1 To x2 step dirx*2
             .writetext chrw((ImgArray(x,y) and 255)+ &h100& *(ImgArray(x+dirx,y) and 255))
          Next
          'line padding
          if pad and 1 then .writetext  chrw(ImgArray(x2,y))
          if pad >1 then .writetext  chrw(0)
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
Set X = (New ImgClass)(fn,w*2,h*2,1,32,0,0)

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

