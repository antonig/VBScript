option explicit

' Raster graphics class in VBSCRIPT by Antoni Gual
'--------------------------------------------
' An array keeps the image allowing to set pixels, draw lines and boxes in it. 
' at class destroy a bmp file is saved to disk and the default viewer is called
' The class can work with 8 and 24 bit bmp. With 8 bit uses a built-in palette or can import user's


'Declaration : 
' Set MyObj = (New ImgClass)(name,width,hight,bits_per_pixel,palette_array)
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
'    GetRGB(r,g,b) Gets a color-code depending of the color depth : if 8bits : nearest color
'
'to do:
'    add clipping to line
'    allow loading of palettes <256
'    background image loading
'    dithering
'    ellipse, arc
'    count colors in 24 bit mode 
'    RLE for 8 bit

Class ImgClass
  Private ImgL,ImgH,ImgDepth,bytepix
  public ImgArray()  'rgb in 24 bit mode, indexes to palette in 8 bits
  private filename   
  Public Palette
  
  Public Property Let depth (x) 
     if depth=8 or depth =24 then 
        Imgdepth=depth
     else 
        Imgdepth=8
     end if
      bytepix=imgdepth/8
  end property        

  Public Property Let Pixel (x,y,color)
    If (x<ImgL) And (x>=0) And (y<ImgH) And (y>=0) Then 'Clipping
      Select Case ImgDepth
      Case 24
          ImgArray(x,y)=Color
      Case 8 
          ImgArray(x,y)=Color Mod 256
      Case Else
          WScript.Echo "ColorDepth unknown : " & ImgDepth & " bits"
      End Select
    End If
  End Property

  Public Property Get Pixel (x,y)
    If (x<ImgL) And (x>=0) And (y<ImgH) And (y>=0) Then
      Pixel=ImgArray(x,y)
    End If
  End Property
  
  Public Property Get ImgWidth ()
    ImgWidth=ImgL-1
  End Property
  
  Public Property Get ImgHeight ()
    ImgHeight=ImgH-1
  End Property     

  Public Default Function Init(name,w,h,dep,pal)
     ImgL=w
     ImgH=h
     redim imgArray(ImgL-1,ImgH-1)
     filename=name
     depth =dep
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
    
  Private Sub Class_Terminate
    savebmp
    wscript.echo "opening " & filename
    CreateObject("Shell.Application").ShellExecute filename
  End Sub


  Sub WriteLong(ByRef Fic,ByVal k)
    Dim x
    For x=1 To 4
        Fic.Write chr(k and &hFF)
        k=k\256
    Next
  End Sub

  Public Sub SaveBMP
    'Save the picture to a bmp file
    Const ForReading = 1 
    Const ForWriting = 2
    Const ForAppending = 8
    Dim Fic
    Dim i,r,g,b
    Dim k,x,y,Pal,padding
    
    pal=-(bytepix=1)
    
    Set Fic = WScript.CreateObject("scripting.Filesystemobject").OpenTextFile(filename, ForWriting, True)
    if fic is nothing then wscript.echo "error creating file" & filename :wscript.quit
    
    dim bms:bms=ImgH* 4*(((ImgL*bytepix)+3)\4)  'bitmap size including padding
    dim pals:pals=(ubound(palette)+1)*4*pal
    
    'FileHeader
    Fic.Write "BM" 'Type
    WriteLong Fic, 14+40+ pals + bms    'Size of entire file in bytes
    fic.write string(4,0)
    WriteLong Fic,54+pals   '2 words: offset of BITMAPFILEHEADER (access to the beginning of the bitmap) 54=14+40 (fileheader+infoheader)

    'InfoHeader
    WriteLong Fic,40    'Size of Info Header(40 bytes)
    WriteLong Fic,ImgL
    WriteLong Fic,ImgH
    Fic.Write chr(1) & chr(0) 'Planes : 1
    Fic.Write chr(ImgDepth) & chr(0) 'Bitcount : 1,4,8,16,24,32 = bitsperpixel
    fic.write string(8,0)&chr(&Hec)&chr(4)& string(2,0)&chr(&Hec)&chr(4)& string(2,0)& string(8,0) 
    
    'palette
    If Pal=1 Then
      For i=0 to 255
        writelong fic ,Palette(i)
      Next
    End If
    
    'write bitmap
    dim xx:xx=(ImgL*bytepix) mod 4
    if xx<>0 then padding=Space(4-xx) else padding=""
    Select Case ImgDepth
    Case 24
      For y=ImgH-1 to 0 step-1  'Origin of bitmap: bottom left
        For x=0 To ImgL-1
         'writelong fic, Pixel(x,y) 
          k=ImgArray(x,y)    
          Fic.Write chrb(k and &hff)
          k=k\256
          Fic.Write chrb(k and &hff)
          k=k\256
          Fic.Write chrb(k and &hff)
        Next
        Fic.Write padding
      Next
    Case 8
      For y=ImgH-1 to 0 step-1
        For x=0 To ImgL-1
            Fic.Write chr(ImgArray(x,y) and &hff)
        Next
        Fic.Write padding
      Next
    Case Else
        WScript.Echo "ColorDepth unknown : " & ImgDepth & " bits"
    End Select
    
    Fic.Close
    Set Fic=Nothing
      
  End Sub
    
  public Sub line(x0,y0, x1,y1,c)
    Dim x,y,xf,yf,dx,dy,sx,sy,err,err2
      x =x0    : y =y0
    xf=x1     : yf=y1
    dx=Abs(xf-x) : dy=Abs(yf-y)
    If x<xf Then sx=+1: Else sx=-1
    If y<yf Then sy=+1: Else sy=-1
    err=dx-dy
    Do
      pixel(x,y)=c   'the pixel property does the clipping... slow!
      If x=xf And y=yf Then Exit Do
      err2=err+err
      If err2>-dy Then err=err-dy: x=x+sx
      If err2< dx Then err=err+dx: y=y+sy
    Loop
  End Sub 'draw_line 

  public sub fbox (x0,y0, w,h,byval c)  
    dim i,j 
    dim x1,y1,x2,y2
    'clipping
    if x0<0 then x1=0 else x1=x0 
    if y0<0 then y1=0 else y1=y0 
    if x0+w>=ImgL then x2=Imgl-1 else x2=x0+w 
    if y0+h>=ImgH then y2=ImgH-1 else y2=y0+h 
      for i = x1 to x2  
        for j = y1 to y2 
          ImgArray(i,j)=c
        next
      next
  end sub
  public sub circle(x0,y0,r,c)
    dim x,y,err 
    x=r:y=0:err=0
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


Dim X,i,j
const sq=20 

Set X = (New ImgClass)("c:\temp\red_on_black2.bmp",sq*16,sq*16,8,0)

'display palette
for i= 0 to 15
  for j=0 to 15
    x.fbox i*sq,j*sq,sq,sq,(j*16+i)
  next
next
'do a couple of lines
x.line 0,0,x.imgwidth,x.imgheight,15
x.line 0,x.imgheight,x.imgwidth,0,0
x.circle 160,160,100,128
Set X = Nothing
