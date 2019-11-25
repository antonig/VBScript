'prueba de funcion que crea un formulario html y devuelve valores en diccionario
'pdte:
'ok  checkboxes
'ok  alinear
'    detectar claves repetidas
'ok  tamaño letra funcion de pantalla?
'    listas
'    
'http://alistapart.com/article/prettyaccessibleforms/

a=array(array("text","myname","nombre","pepe"),_
        array("text","myaddr","direccion","Viladomat, 210"),_
        array("radio","gender","male","&female","other"),_
	array("check","asignatura","&mates","fisica","historia"))

set d=getuserinput(a,"escribir valores", 320,300 )
if not d is Nothing  then
  for each i in d.keys
    wscript.echo  i & " "& d(i) 
  next
else
  wscript.echo "Aborted"
end if


Function GetUserInput( myPrompt,title,w,h )
    Set objIE = CreateObject( "InternetExplorer.Application" )
    ' Specify some of the IE window's settings
    objIE.Navigate "about:blank"
    objIE.Document.title = title
    objIE.ToolBar        = False
    objIE.Resizable      = False
    objIE.StatusBar      = False
    objIE.Width          = w
    objIE.Height         = h

    ' Center the dialog window on the screen
    With objIE.Document.parentWindow.screen
        objIE.Left = (.availWidth  - objIE.Width ) \ 2
        objIE.Top  = (.availHeight - objIE.Height) \ 2
    End With

    ' Wait till IE is ready
    Do While objIE.Busy
        WScript.Sleep 200
    Loop
    
    ' Build HTML code of the form and dictionnary with keys from parsing the array passed as argument
    Dim d:Set d = CreateObject ("Scripting.Dictionary")
    
    s = "<table width=""100%"" cols=""2"" style=""font-size:5vw"" id=""form"" >"
    for each a in myprompt 
      select case a(0)
      case "text"
        d.Add a(1),""  
        s=s & vbcrlf & "<tr><td>" & a(2) &"</td><td align=""right""><input style=""font-size:5vw"" type=""text""   name=""" & a(1) & """ value=""" & _
           a(3) & """ />  </tr>" & vbcrlf
      case "radio"
        d.Add a(1),"" 
        s=s & vbcrlf &"<tr><td align=""center"" colspan=""2""><fieldset ><legend>"&a(1)&"</legend> " & vbCrlf
        for i=2 to ubound(a)
          if left(a(i),1)="&" then 
              chk="checked":j=mid(a(i),2) 
          else 
              j=a(i):chk=""
          end if 
          s=s& vbcrlf & "<input type=""radio""  Name="""& a(1) & """ value=""" & j &""" " & chk &"><label> " & j &"</label>"  '    &" <br /> "
        next
        s=s & vbcrlf & "</div></fieldset></td></tr>"
      case "check"  
        s=s & vbcrlf &"<tr><td colspan=""2"" align=""center""><fieldset><legend>"&a(1)&"</legend>" & vbCrlf    
        for i=2 to ubound(a)
          if left(a(i),1)="&" then 
              chk="checked":j=mid(a(i),2) 
          else 
              j=a(i):chk=""
          end if 
          d.Add j,"" 
          s=s& vbcrlf & "<input type=""checkbox""  Name="""& j &""" value=""" & j &""" " & chk &">" & j 
        next
        s=s & vbcrlf & "</td></tr>" 
      end select
    next        
    s=s & vbcrlf & "<tr><td align=""center"" colspan=""2""><input type=""hidden"" id=""OK"" name=""OK"" value=""0"">" _
         & "<input type=""submit"" value="" OK "" OnClick=""VBScript:OK.value=1"">" _
         & "</tr></td></table>"
  
    wscript.echo s
    'make it visible
    objIE.Document.body.innerHTML =s
    objIE.Document.body.style.overflow = "auto"
    objIE.Visible = True
    
    'wait for ok button
    On Error Resume Next
    Do 
      WScript.Sleep 200
    Loop While objIE.Document.all.OK.value = 0 
    'User has cancelled
    If Err then objIE.Quit:set getuserinput=nothing: Set objIE = Nothing: exit function
    On Error Goto 0
    
    'read fields to dictionnary 
    for each a in d.keys
      set r= objIE.Document.getelementsbyName(a)
      select case r(0).type 
      case "radio"
        for b=0 to r.length-1
          if r(b).checked then n=r(b).value
        next
        d(a)=n
      case "checkbox"
        for b=0 to r.length-1
          n= r(b).checked 
          d(a)=n
        next
      case "text"
        n=r(0).value
        d(a)=n              
      case else
      end select	 
    next  
    set GetUserInput = d
    objIE.Quit
    Set objIE = Nothing
End Function