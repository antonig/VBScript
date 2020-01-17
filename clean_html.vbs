option explicit
'---------------------------------
' utiliza regex para limpiar html y dejar solo texto
'------------------------------------

dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
dim ows: Set ows = CreateObject("Wscript.shell")
 '   Const fsoForReading = 1
  '  Const fsoForWriting = 2
 
    
dim f:f=selecfile
 
dim s:s=loadstringfromfile(f,"",1)
dim a,b
re_replace s,"<!--[\s\S]+?-->",""
re_replace s,"<a href[\s\S]+?</a>",""
re_replace s,"<picture[\s\S]+?</picture>",""
re_replace s,"<script[\s\S]+?</script>",""
re_replace s,"<svg[\s\S]+?</svg>",""
re_replace s,"<nav[\s\S]+?</nav>",""
re_replace s,"<figure[\s\S]+?</figure>",""
re_replace s,"<img[\s\S]+?/>",""
re_replace s,"<link[\s\S]+?>",""
re_replace s,"<meta[\s\S]+?>",""
re_replace s,"<form[\s\S]+?</form>",""
re_replace s,"<input[\s\S]+?/>",""
re_replace s,"\s+?(\n)","$1"  'borra blancos fin linea
re_replace s,"\n*(\n)","$1"   'borra lineas vacias 
blocdenotas s,0,".html",vbcrlf,1 
wscript.quit(0)





sub re_replace(a,byval p,byval r)
 with New RegExp
  .pattern=p
  .global=True
  .multiline=true
  a =.replace (a,r)
 end with
end sub


sub blocdenotas( byref a,cnt,nom,sep,utf)
'escribe texto ascii o utf-8 a archivo nom y abre bloc de notas ara visualizarlo
's   cadena texto o array valores
'cnt longitud array
'nom nombre archivo si "" se crea nombre.extension, si ".xxx" se usa extension
'sep si s es array, cadena a usar como separador, si s es cadena se ignora
'utf 1 si Charset v a ser utf8, 0 si ascii
 if isarray(a) then 
    redim preserve a(cnt)
    s=join(a,sep)
    erase a    
 else 
   s=a
 end if   
 if nom="" then 
    nom=fso.gettempname
 elseif left(nom,1)="." then 
    nom=replace(fso.gettempname,".tmp",nom)
 end if
 With CreateObject("ADODB.Stream")
     .Open
     if utf then .CharSet = "utf-8"   
     .WriteText s
     .SaveToFile nom, 2
 End With
 ows.run "notepad " & nom,,0
end sub


function scriptpath()
  'get the path of this script 
  scriptpath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
end function	

function selecfile
'si archivo pasado como primer parametro exise, devuelve su nombre
'si no existe o no se han pasado parameros abre selector de archivos de windows
 if (Wscript.Arguments.count=1) then  
   dim f:f=Wscript.Arguments(0)
   if instr(f,"\")=0 then f=scriptpath & f     
   if fso.FileExists(f) Then
    Selecfile=f
    exit function
  end if
end if	
with  CreateObject("Wscript.shell").Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();"&_
   "new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);"&_
   "close();resizeTo(0,0);</script>""")
    Selecfile = .StdOut.ReadLine
 end with   
 
end function

Function LoadStringFromFile(filename,sp,utf)
    'lee texto de archivo, si se especifica p devuelve array cortado en separadores
    'filename nombre archivo si no se da path se busca en carpeta del script
    'sp       si es "" se devuelve string, si no se devuelve array usando split
    'utf      si es 1 se usa charset utf-8 si 0 se usa ascii 
    dim s
    if instr(filename,"\")=0 then filename=scriptpath & filename 
    With CreateObject("ADODB.Stream")
     .Open
     if utf then .CharSet = "utf-8"
     .loadfromfile filename
     s= .readtext
     .Close
    end with
    if len(sp) then 
       LoadStringFromFile= split(s,sp)
    else
       LoadStringFromFile= s
    end if
End Function
