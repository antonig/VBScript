option explicit
'-----------------------------------------------------
'SRT time offset by Antoni Gual Via 1-2020
'----------------------------------------------------
'Scans a srt video subtitle file and generates an output file
' with all timings offset by the same amount of seconds.
' (useful when the video intro is shortened after generating the SRT) 
' The output file gets a random name and its opened in Notepad. 
' Rename and save it in the folder where it must go.
'----------------------------------------------------
dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
dim ows: Set ows = CreateObject("Wscript.shell")
    Const fsoForReading = 1
    Const fsoForWriting = 2
    
    
dim f:f=selecfile()
dim p:p=scriptpath
dim a: a=timeserial(0,0,Cint(inputbox("Offset in sec.","srt file time editor",0)))
dim s:s=loadstringfromfile( f)
dim re:set re= new regexp

re.pattern="(\d{2}:\d{2}:\d{2})(.+?)(\d{2}:\d{2}:\d{2})(.+)"


dim i:for i=0 to ubound(s)
   if re.test(s(i)) then
    dim m:set m= re.execute(s(i))
    dim sm:set sm=m(0).submatches
    'wscript.echo istime(m(0)),istime(m(1))
     s(i)=formatdatetime(timevalue(sm(0))+a,3) & sm(1)& Formatdatetime(timevalue(sm(2))+a,3)& sm(3) 
  end if
next
blocdenotas s,ubound(s),"" 
wscript.quit(0)

sub blocdenotas( byref a,cnt, nom)
'escribe texto a en logfile nom y abre bloc de notas ara visualizarlo
 redim preserve a(cnt)
 dim s: if isarray(a) then s=join(a,vbcrlf)  else s=a
 if nom="" then nom=fso.gettempname
 Dim LogFile:Set LogFile = fso.CreateTextFile(nom, true)
 logfile.write s
 LogFile.Close
 ows.run "notepad "&nom,,0
 erase a
end sub

function scriptpath()
  scriptpath = FSO.getparentfoldername(fso.getfile(Wscript.ScriptFullName))&"\"
end function	

function selecfile
 if Wscript.Arguments.count=1 then
    Selecfile=Wscript.arguments(0)
 else	
 dim oexec
 Set oExec=ows.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();"&_
   "new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);"&_
   "close();resizeTo(0,0);</script>""")
 Selecfile = oExec.StdOut.ReadLine
 end if
end function

Function LoadStringFromFile(filename)
    Dim   f,s
    Set f = fso.OpenTextFile(filename, fsoForReading)
    s= f.ReadAll
    f.Close
    loadstringfromfile=split(s,vbcrlf)
End Function

function IsTime (str)  'not used!
  if len(str) = 0 then
    IsTime = false
  else
    On Error Resume Next
    TimeValue(str)
    if Err.number = 0 then
      IsTime = true
    else
      IsTime = false
    end if
    On Error GoTo 0
  end if
end function