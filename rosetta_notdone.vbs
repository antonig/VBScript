Set http= CreateObject("WinHttp.WinHttpRequest.5.1")
Set oDic = WScript.CreateObject("scripting.dictionary")
Set ows = WScript.CreateObject("Wscript.Shell")
Set myfso = WScript.CreateObject("Scripting.Filesystemobject")


start="https://rosettacode.org"
Const lang="VBScript"
Dim oHF 
print "getting official tasks for " & lang 
gettaskslist "about:/wiki/Category:Programming_Tasks" ,True
print odic.Count
print "getting draft  tasks"
gettaskslist "about:/wiki/Category:Draft_Programming_Tasks",True
print "total tasks " & odic.Count
print "Removing tasks already done in " & lang
gettaskslist "about:/wiki/Category:"&lang,False

salida_notepad odic,"VBS No_en_rosetta.txt",1
print "total tasks  not in " & lang & " " &odic.Count & vbcrlf
pause
WScript.Quit(1)

Sub pause() wscript.stdout.write  "Press Enter to Continue":wscript.stdin.readline: End Sub
  
  Sub print(s): 
    On Error Resume Next
    WScript.stdout.WriteLine (s)  
    If  err= &h80070006& Then WScript.echo " Please run this script with CScript": WScript.quit
  End Sub 
  
  Function getpage(name)
    Set oHF=Nothing
    Set oHF = CreateObject("HTMLFILE")
    http.open "GET",name,False  ''synchronous!
    http.send 
    oHF.write "<html><body></body></html>"
    oHF.body.innerHTML = http.responsetext 
    Set getpage=Nothing
  End Function
  
  
  Sub gettaskslist(b,build)
    nextpage=b
    While nextpage <>""
      
      nextpage=Replace(nextpage,"about:", start) 
      'print nextpage
      getpage(nextpage)
      Set xtoc = oHF.getElementbyId("mw-pages")
      nextpage=""
      For Each ch In xtoc.children
        If  ch.innertext= "next page" Then 
          nextpage=ch.attributes("href").value
          ': print nextpage
        ElseIf ch.attributes("class").value="mw-content-ltr" Then
          Set ytoc=ch.children(0) 
          'print ytoc.attributes("class").value  '"mw-category mw-category-columns"
          Exit For
        End If   
      Next
      For Each ch1 In ytoc.children 'mw-category-group
        'print ">" &ch1.children(0).innertext &"<"
        For Each ch2 In ch1.children(1).children '"mw_category_group".ul
          Set ch=ch2.children(0)
          If build Then
            odic.Add ch.innertext , ch.attributes("href").value
          Else    
            If odic.Exists(ch.innertext) Then odic.Remove ch.innertext
          End If   
          'print ch.innertext , ch.attributes("href").value
        Next 
      Next
    Wend  
  End Sub
  
  'a bloc de notas si es array o diccionario combina, si no se da nombre 
  Sub salida_notepad(stuff,tname,espera)
    Dim st
    If TypeName(stuff)="Dictionary" Then
      ReDim a(stuff.count-1)
      Dim cnt,i
      For Each i In stuff.keys
        a(cnt)= i& vbTab & stuff(i)
        cnt=cnt+1 
      Next 
      st=Join(a,vbCrLf)
      Erase a
    Else
      If IsArray(stuff) Then st=join(stuff,vbcrlf) Else st=stuff
    End if   
    With myFSO.OpenTextFile(tname, 2, True)
      .WriteLine(St) 
      .Close
    End With	
    ows.run  "notepad.exe " & tname,,espera
    If espera And myfso.fileexists(f) Then myfso.deletefile(tname)
  End Sub
  
  
  
  
  
  
  
  
  