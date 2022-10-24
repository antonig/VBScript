'LZW compress - uncompress  functions in plain VBScript  By Antoni gual 2022
'----------------------------------------------------------------------------
'I just made several times faster an idea by Rabbit found here
'https://bytes.com/topic/access/insights/962727-lzw-compression-algorithm-vbscript
'------------------------------------------------------------------------------
'Notes: 
'It's slow, compresses at a rate of 5 secs/Mb it uses only the usual VBScript resources, fileObject, 
'   Dictionary and ADODB.Stream
'No auxiliar dll to register
'
'The compress code use an UTF-16 stream so it adds a BOM at the start of the file
'making the compressed file longer by 2 bytes. I' did'nt care to remove it..
'
'The decompression is much faster than the compression. This is because
'compression requires a dictionnary while decompression works with a simple array of strings
'
'All the work is made with both files in memory: input file is read with a readall
'  while output file is written to a memory stream. 
'
'Do not test with an already compressed file .jpg .gif .png .docx .xlsx ....
'
' Tested in WinXP SP3, Windows 7 Professional 32 and Windows10. Run it with cscript
'------------------------------------------------------------------------------

Option Explicit
Const numchars=255  'good for binary files
Const maxcod=32767  'maximum 32767 or chrw will complain!
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Function LZWCompress(strPath,wpath)
  Dim oDict, strNext, strCurrent, intMaxCode, i,z,si,ii,ofs,ofread
  Set oDict = CreateObject("Scripting.Dictionary")
  With WScript.CreateObject("ADODB.Stream") 
    .type= 2
    .charset= "x-ansi"
    .open
    .loadfromfile strpath
    si=.readtext
    .close
   End with  
 With WScript.CreateObject("ADODB.Stream") 
    .type=2 '2text  1binary
    .charset= "utf-16le"
    .open  
    
    intMaxCode = numchars
    For i = 0 To numchars
      oDict.Add Chr(i), i
    Next
    
    strCurrent = Left(si,1)
    For ii=2 To Len(si)
      strNext = Mid(si,ii,1)
      
      If oDict.Exists(strCurrent & strNext) Then
        strCurrent = strCurrent & strNext
      Else
        .writetext ChrW(oDict.Item(strCurrent)) 
        intMaxCode = intMaxCode + 1
        oDict.Add strCurrent & strNext, intMaxCode
        strCurrent = strNext
        
        If intMaxCode >= maxcod Then
          oDict.RemoveAll
          intMaxCode = numchars
          For i = 0 To numchars
            oDict.Add Chr(i), i
          Next
        End If
      End If
    next
    .writetext ChrW(oDict.Item(strCurrent))
    .savetofile wpath,2 'adsavecreateoverwrite
    .Close
    
  End with
  
  Set oDict = Nothing
End Function

Function lzwUncompress(strpath, wpath)
  Dim intNext, intCurrent, intMaxCode, i,ss,istr
  Set istr= WScript.CreateObject("ADODB.Stream")
    istr.type=2
    istr.charset="UTF-16LE"
    istr.Open
    istr.loadfromfile strpath
    istr.Position=2
  With WScript.CreateObject("ADODB.Stream")
    .type=2 '2text  1binary
    .charset= "x-ansi"
    .open          
    reDim dict(maxcod)
    intMaxCode = numchars
    For i = 0 To numchars : dict(i)= Chr(i) :  Next
      'strNext = Left(si,1)
      intCurrent=ascw(istr.readtext(1))
      
      While Not istr.EOS
        ss=dict(intCurrent)
        .Writetext ss
        intMaxCode = intMaxCode + 1
        Dim x:x=istr.ReadText(1)
       intNext=ascw(x) 
        If intNext<intMaxCode Then
          dict(intMaxCode)=ss & Left(dict(intNext), 1)
        Else
          dict(intMaxCode)=ss & Left(ss, 1) 
        End If
        If intMaxCode = maxcod Then  intMaxCode = numchars
        intCurrent = intNext
      wend
      .Writetext Dict(intCurrent) 
      .savetofile wpath,2 'adsavecreateoverwrite
      .Close
    End With
    istr.close
    Set istr=nothing
End function

sub print(s): 
   On Error Resume Next
   WScript.stdout.WriteLine (s) : 
   if  err= &h80070006& then WScript.Echo " Please run the script with CScript ":WScript.quit
End sub 

function filesize(fn) filesize= CreateObject("Scripting.FileSystemObject").GetFile(fn).Size: end function
function fileexists(fn) fileexists= CreateObject("Scripting.FileSystemObject").FileExists(fn): end function
'-----------------------------------------------
Dim t,t1,x,a,b,c

If WScript.Arguments.Count >0 Then
  a=WScript.Arguments(0)
else  
 a="unixdict.txt" ' The file To test with, shoudl be in the script folder
 a=Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))&a
End If
b=a & ".lzw"
c=a & ".unc"
If Not fileexists(a) Then print "File " &a & " not found":WScript.quit
 
t1=Timer
print "Compressing " &a
LZWCompress a,b
print Timer-t1 & " Seconds"
print "Size reduction: "& FormatNumber(100-(100*filesize(b) / filesize(a)),2) &"%"

print "decompressing " &b 
t=timer
LZWUncompress b,c
print Timer-t & " seconds"
print Timer-t1 & " seconds Total"
 
x= "cmd /k fc """ & a & """ """ & c & """ /N  &pause & exit " 
CreateObject("WScript.Shell").run x,,true
wscript.quit 1  