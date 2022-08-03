option explicit
'Displays an html Form  and returns the user selections in a dictionary
'http://alistapart.com/article/prettyaccessibleforms/
'My contribution: 
 '  It can get the user form from a file or use the included MakeHtmlForm function
'   MakeHtmlForm function builds an html form without input validation from a simple array of arrays

'MakeHTMLForm argument format:
   'a single argument: an array of arrays 
   'each sub-array defines 1 field in a new line
    
   '"text",Caption, default  
   '"radio",GroupCaption,list of Itemcaptions  (the one prefixed by &  is the default)
   '"check",GroupCaption, list of ItemCaptions (the ones prefixed by & are preselected)
   '   
   ' the function returns the body of the html form

'Getuserinput   
   'arguments:   
     'myform is the body part of the form
     'title is window title 
     'w h width height of the window

    'the function returns a dictionnary  with 
       'nothing if dialog cancelled  
       'for text fields returns the caption as the key and the user input (or the default)
       'for radio button fields returns GroupCaption as the key and the selected item caption as tha value
       'for check box fields returns the ItemCaptions as the keys and true or false as the values
       
'TO DO:
'ok  checkboxes
'ok  alinear
'ok  array de strings y join para generar html
'ok  tamaño letra funcion de pantalla?
'ok  opcion traer html de archivo. Lo cargo a variable, no uso IE.Navigate para saltarme seguridad del IE11
'    detectar claves repetidas
'    drop down lists (selects)
'    multiline textinputs
'    captions (color!)     
'    validacion 
'    reset button
'    reacciones



Function ReadAllTextFile(fname)   'used to read predefined html form
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   on error resume next
     Set f = fso.OpenTextFile(fname, ForReading)
		 if err then readalltextfile=null :exit function
    on error goto 0
   ReadAllTextFile =   f.ReadAll
End Function


Function MakeHtmlForm(Myprompt)
   'converts myprompt (array of arrays) to valid html 

    dim a,s,chk,j,i,sc
    redim s1(1000)
    sc=0:s1(sc) = "<table width=""100px"" cols=""2"" style=""font-size:5vw""  >"
    
    for each a in myprompt 
      select case a(0)
      case "text"
          
        sc=sc+1:s1(sc)="<tr><td>" & a(2) &"</td><td align=""right""><input style=""font-size:5vw"" type=""text""   name=""" & a(1) & """ value=""" & _
           a(3) & """ />  </tr>"  
      case "radio"
        
        sc=sc+1:s1(sc)="<tr><td align=""center"" colspan=""2""><fieldset ><legend>"&a(1)&"</legend> " 
        for i=2 to ubound(a)
          if left(a(i),1)="&" then 
              chk="checked":j=mid(a(i),2) 
          else 
              j=a(i):chk=""
          end if 
          sc=sc+1:s1(sc)= "<input type=""radio""  Name="""& a(1) & """ value=""" & j &""" " & chk &"><label> " & j &"</label>"  
        next
        sc=sc+1:s1(sc)="</div></fieldset></td></tr>"
      case "check"  
        sc=sc+1:s1(sc)="<tr><td colspan=""2"" align=""center""><fieldset><legend>"&a(1)&"</legend>" 
        for i=2 to ubound(a)
          if left(a(i),1)="&" then 
              chk="checked":j=mid(a(i),2) 
          else 
              j=a(i):chk=""
          end if 
           
          sc=sc+1:s1(sc)= "<input type=""checkbox""  Name="""& j &""" value=""" & j &""" " & chk &">" & j 
        next
        sc=sc+1:s1(sc)="</td></tr>" 
      end select
    next        
    sc=sc+1:s1(sc)="<tr><td align=""center"" colspan=""2""><input type=""hidden"" id=""OK"" name=""OK"" value=""0"">"
    sc=sc+1:s1(sc)="<input type=""submit"" value="" OK "" OnClick=""VBScript:OK.value=1"">"
    sc=sc+1:s1(sc)= "</tr></td></table> "
    redim preserve s1(sc+1)
    s=join(s1,vbcrlf)
    wscript.echo ubound(s1),s
    MakeHtmlForm=s
end function

Function GetUserInput( myForm,title,w,h)
   'open an IE windows with user form  
   'user can fill fields. At OK press, VBS parses answers to a dictionnary and kills IE window  	
   'myform is formatted html
   'title is window title ,w h width height in pixels    
   '--------------------------------------------------------- 
    dim objie,n,b,r
    Set objIE = CreateObject( "InternetExplorer.Application" )
    with objIE

    ' Specify some of the IE window's settings
    .Navigate "about:blank"
    .Document.title = title
    .ToolBar        = False
    .Resizable      = False
    .StatusBar      = False
    .width          =w
		.height         =h  

    ' Center the dialog window on the screen
    'With .Document.parentWindow.screen
    '    objIE.Left = (.availWidth  - objIE.Width ) \ 2
    '    objIE.Top  = (.availHeight - objIE.Height) \ 2
    'End With

    ' Wait till IE is ready
    Do While .Busy
        WScript.Sleep 200
    Loop
    
    'make it visible
    .Document.body.innerHTML =myForm
    .Document.body.style.overflow = "auto"
    .Visible = True
    
    'wait for ok button or an abort error
    On Error Resume Next
    Do 
      WScript.Sleep 200
    Loop While .Document.all.OK.value = 0 

    'User has aborted: return nothing
    If Err then .Quit: set GetUserInput=nothing: Set objIE = Nothing: exit function
    On Error Goto 0
    
    ' create a dictionnary with user answers
    dim d:Set d = CreateObject ("Scripting.Dictionary") 
    'read fields to dictionnary ya no va!
    dim f :set f=.Document.getElementsByTagName("Input")  
    for each r in f
       'wscript.echo r.type
      select case r.type 
      case "radio"
          if r.checked then  d.add r.name,r.value 
      case "checkbox"
              d.add r.name,r.checked
      case "text","textarea"
         d.add r.name,r.value             
      case else
         'wscript.echo "Can't add "& r.name
      end select	 
    next  
    set GetUserInput = d
    .Quit
    end with
    Set objIE = Nothing
End Function

dim d,i,s


s=ReadAllTextFile(Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))&"theform.html")
if isnull (s) then
'using the form built by MakeHtmlForm from the definition
	dim a
  wscript.echo "File not found, creating form from built-in structure"
	a=array(array("text","myname","Name","pepe"),_
					array("text","myaddr","Address","Viladomat, 210"),_
					array("radio","Gender","male","&female","other"),_
					array("check","Course","&math","physics","history"),_
				 array("radio","Car","Volvo","Mercedes","Saab","Audi","Other"))
	s=MakeHtmlForm(a)
  
end if

 set d=getUserInput(s,"escribir valores",320,280)  'ya no obedece a
'process results
if not d is Nothing  then
  s="User input:" 
  for each i in d.keys
     s= s & vbcrlf & i & ": "& d(i) 
  next
  wscript.echo s
else
  wscript.echo "Aborted"
end if
wscript.stdout.write vbcrlf&"Press a key": wscript.stdin.read(1)

