'ascii art compressor. Lines should not be longer then 127 bytes!!!

Option Explicit
Dim a,b,o,fso,fn,aa,bb,cc
If WScript.Arguments.Count >0 Then
  aa=WScript.Arguments(0)
else  
 aa="snoopy.txt" ' The file To test with, should be in the script folder
 aa=Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))&aa
End If
bb=aa & ".rle"
cc=aa & ".unc"

if not fileexists(aa) then print "File " & aa & " not found":wscript.quit

Set fso=createobject("Scripting.Filesystemobject")
a=readfile(aa)
o= compress (a)
writefile bb,o
print "Size reduction: "& FormatNumber((100-(100*filesize(bb) / filesize(aa))),2) &"%"
b=decompress (o)
writefile cc,b 
WScript.quit(1)


function filesize(ff) filesize= CreateObject("Scripting.FileSystemObject").GetFile(ff).Size : end Function
function fileexists(fn) fileexists= CreateObject("Scripting.FileSystemObject").FileExists(fn): end function

sub print(s): 
   On Error Resume Next
   WScript.stdout.WriteLine (s) : 
   if  err= &h80070006& then WScript.Echo " Please run the script with CScript":WScript.quit
End sub 

Function makepath(file)
  Dim sfn:  sfn=WScript.ScriptFullName
  sfn= Left(sfn, InStrRev(sfn, "\"))
  makepath= sfn & file
End Function

Function readfile(file)
  Dim a
  With WScript.CreateObject("ADODB.Stream")
    .Open
    .type=2    'texto
    .Charset = "x-ansi"
    .LoadFromFile file
    a=.readtext
    .Close
  End With
  'print "read  "& Len(a)
  readfile=Split(a,vbCrLf)
  If UBound(readfile)=0 Then readfile=Split(a,vbCr)
End Function


Sub Writefile(File, Sdata)  'sin BOM!!
  Dim a
  If IsArray(sdata) Then a=Join(sdata,vbCr) Else a=sdata
  print "writing "&  Len(a)
  With WScript.CreateObject("ADODB.Stream")
    .Open
    .type=2    'texto
    .Charset = "x-ansi"
    .writetext a
    .SaveToFile file ,2 'adSaveCreateOverWrite
    .Close
  End With
End Sub



'codificamos caracter +128, repeticion+32
Function compress (a)
  ReDim b(UBound(a))

  Dim i,j,ch,lch,cnt,s
  For j=0 To UBound(a)
    s=RTrim(a(j))
    If Len(s)>0 then
    cnt=1
    ch=Asc(mid(s,1,1))
    For i=2 To len(s)
      lch=ch
      ch=Asc(mid(s,i,1))
      If ch=lch Then  
        cnt=cnt+1 
        If i= Len(s) Then  B(j)=b(j) & chr(128+ch)  & Chr(32+cnt)  
      ElseIf ch<>lch Then 
        B(j)=b(j) & chr(128+lch)
        If cnt>1 Then B(j)=b(j) & chr(32+cnt):cnt=1 'solo cnt si mas de 1
        If i= Len(s) Then  B(j)=b(j) & chr(128+ch)
      End If 
     Next
     End if
   
  Next
  compress=b
End Function

Function decompress(a)
  Dim i,j,c,lc,w
  ReDim b(UBound(a))
  For i=0 To UBound(a)
    'cada linea
    lc=0
    If Len(a(i))>0 then
    c=Asc(Mid(a(i),1,1))
    For j=2To Len(a(i))
      lc=c:w=0
      c=Asc(Mid(a(i),j,1))
      If c<128 Then 
         B(i)=b(i) & string(c-32,Chr(lc-128)) 
         w=1
      ElseIf c>128  And lc>128 Then 
         B(i)=b(i) & Chr(lc-128)
         If j= Len(a(i)) Then  B(i)=b(i) & Chr(c-128)       
         w=1
      End if   
    Next 
    If c>128 And w=0 Then  B(i)=b(i) & Chr(c-128)  
    End if 
  Next
  decompress=b
End Function
