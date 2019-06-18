option explicit

dim arr:arr=array(2,-7,0,14,4,30,3,-8,7,7)
dim oconncsv:Set oConnCsv = CreateObject("ADODB.Connection")

dim osh: set osh=CreateObject("Wscript.Shell")

dim mypath: mypath= osh.ExpandEnvironmentStrings("%userprofile%")&"\Desktop"
dim myfile: myfile="Export_Delta.csv"

dim qry:qry="SELECT * FROM " & myfile & " WHERE DEV_ID=" & 100

dim rs:set rs= getcsv32(mypath,qry)
wscript.echo(view_rs(rs,arr))
 
wscript.quit(0)

'---------------------------------------------------

function getcsv32(path,query)
'requiere vbs 32 bits
'path sin nombre de archivo
'query SELECT con nombre tabla= nombre archivo csv
'devuelve recordset
dim orscsv:Set oRsCsv = CreateObject("ADODB.Recordset")
dim connstring
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001	
Const provi="Microsoft.Jet.OLEDB.4.0"
	connstring="Provider=" & provi   &";" & _
			  "Data Source=""" & path  & """;" & _
			  "Extended Properties=""text;HDR=YES;FMT=Delimited"""
	wscript.echo connstring
  'on error resume next

   if err then wscript.quit(1)
   wscript.echo query
   oRsCsv.Open query,connstring, adOpenStatic, adLockOptimistic, adCmdText
   if err then wscript.quit(2)
   'on error goto 0
   wscript.echo "consulta efectuada. registros: "  & oRsCsv.RecordCount
   set getcsv32=orscsv
end function


function view_rs(r, a) 
'r es un recordset obtenido de consulta
'a array que alterna numeros de columna (base 0) y espacios(negativo alinea derecha)
dim s,i,t,l,c
redim l(r.recordcount+1)
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



