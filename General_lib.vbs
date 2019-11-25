public ows       : Set oWs = CreateObject("WScript.Shell")
public fso       : Set fso = CreateObject("Scripting.FileSystemObject")
'--------------------------------------------------------------
'to have this lib included include this sub in your main file and call it from the first lines with the name of this file
'Sub includeFile(fSpec)
'    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
'End Sub
'---------------------------------------------------------
function selecfile
 'uses hta to run a windows file selector dialog
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
'-----------------------------------------------------------------
function gettable32(path,query)
'devuelve recordset desconectado de tabla excel o csv
'si path incluye nombre archivo xls query debe tener nombre_hoja como nombre tabla
'si path no incluye nombre archivo query debe tener nombre_archivo.csv como nombre tabla

dim oConncsv:Set oRsCsv = CreateObject("ADODB.Connection")
dim orscsv:Set oRsCsv = CreateObject("ADODB.Recordset")
dim connstring 
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001	
Const provi="Microsoft.Jet.OLEDB.4.0"
if lcase(right(path,3))="xls" then
    connstring="Provider=" & provi   &";" & _
       "Data Source=""" & path & _
       """; Extended Properties=""Excel 8.0;HDR=Yes;"";" 
else
    connstring="Provider=" & provi   &";" & _
      "Data Source=""" & path  & """;" & _
      "Extended Properties=""text;HDR=YES;FMT=Delimited"""
end if
	wscript.echo connstring
  on error resume next
   oConnCsv.Open connstring
    if err then terminar "No puede conectarse a CSV " & path
     wscript.echo query
   oRsCsv.Open query,oConnCsv, adOpenStatic, adLockOptimistic, adCmdText
   if err then terminar "No puede hacerse consulta CSV " & query
  on error goto 0
   wscript.echo "consulta efectuada. registros: "  & oRsCsv.RecordCount
   oRsCsv.Activeconnection=nothing  
   set gettable32=orscsv
   set oConnCsv=Nothing
   set oRsCsv=Nothing
 end function
 '-------------------------------------------------
 function view_rs(r, a ) 
 'devuelve recordset  en string
dim s,i,t,l,c
redim l(r.recordcount+1)
'r es un recordset obtenido de consulta
'a array que alterna numeros de columna (base 0) y espacios(negativo alinea derecha)
	with r.Fields
    s=""
    for i=0 to ubound(a) step 2
      t=a(i+1) 
      if t<0 then  
        S= s& right(space(-t)&.Item(a(i)).name & " ",-t )
      else
        S= s& left(.Item(a(i)).Name & space(t),t)
      end if        
    next
    l(0)= s
    c=1
    Do Until r.EOF
      s="" 
      for i=0 to ubound(a) step 2
        t=a(i+1)
        if t<0 then  
          S= s& right(space(-t)&.Item(a(i))&" ",-t)
        else
          S= s& left(.Item(a(i))& space(t),t)
        end if  
      next  
      l(c)=s
      r.MoveNext:c=c+1
    Loop
    end with
    view_rs=join(l,vbcrlf)
end function

'----------------------------------------------------------------------

'-----------------------------------------------
sub blocdenotas( byref a, nom)
'admite array de strings o string
'escribe texto a en logfile nom y abre bloc de notas ara visualizarlo
 dim s: if isarray(a) then s=join(a,vbcrlf)  else s=a
 if nom="" then nom=fso.gettempname
 Dim LogFile:Set LogFile = fso.CreateTextFile(nom, true)
 logfile.write s
 LogFile.Close
 ows.run "notepad "&nom,,0
 erase a

end sub


'--------------------------------------------------------------------
'ASEGURAR HOST CSCRIPT O WSCRIPT Y 32 BITS
'------------------------------------------------------------------
'Asegura host restarts your script with the corrent settings to ensure it runs 
'with the correct version of the vbs engine (console/windows or 32-64 bits)

function EsWin64   'devuelve 1 si es Windows 64 bits
    EsWin64= (GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth = 64 )
end function

function nomhost (mode) ' mode es cmd o win
   if ucase(mode)="CMD" then nomhost= "CSCRIPT.EXE" else nomhost= "WSCRIPT.EXE" 
end function

function eshost(mode)  ' mode es cmd o win
     eshost= (instrrev(ucase(WScript.FullName),nomhost(mode))<>0)
end function

function esbits(bits) 'bits es "32" (forzar 32bits) o ""(los bits que tenga el S.O.) 
    esbits=(instr (ucase(WScript.FullName),nomdir(bits))<>0)
end function

function nomdir(bits) 'devuelve carpeta Sistema para Win32 o Win64
     if bits=32 and eswin64 then nomdir="\SYSWOW64\" else nomdir="\SYSTEM32\"
end function

sub AseguraHost(mode,bits) ' mode "cmd" o "win"  bits "32" o ""(indiferente)
  Dim oProcEnv : Set oProcEnv = oWs.Environment("Process") 
  If (EsWin64 and not esbits(bits)) or not eshost(mode) Then
    Dim sArg, Arg
    If Not WScript.Arguments.Count = 0 Then
      For Each Arg In Wscript.Arguments 
        sArg = sArg & " " & """" & Arg & """"
      Next
    end if 
    Dim sCmd : sCmd = """" &  oProcEnv("windir") & nomdir(bits) & nomhost(mode) & """ " & """" & _
      WScript.ScriptFullName & """ " & sarg
    oWs.Run sCmd
    WScript.Quit
  End If
end sub

'------------------------------------------------------
sub isservicerunning (servicename)
dim flag
'Set wmi = GetObject("winmgmts://./root/cimv2")
on error resume next
flag = (GetObject("winmgmts://./root/cimv2").Get("Win32_Service.Name='" & serviceName & "'").Started)
if err then terminaerror err, "isservicerunning"
on error goto 0
if flag=0 then terminaerror 101, "isservicerunning"
end sub
'------------------------------------

'--------------------------------------------------------

