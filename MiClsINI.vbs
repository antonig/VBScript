Option explicit
' funciona !
' mas rapido solo claves a recordset para ordenar
' enumeracione todas las claves de una seccion?
' como no perder comentarios?
' variable de estado : cargado, sucio
' ahora es sensible a caja, deberia no serlo?

dim mini
set mini=new Clsini
mini.OpenINIFile(Scriptpath()&"sample.ini")
wscript.echo mini.GetINIValue("Sec2","value1")
wscript.echo mini.GetINIValue("pepe","juan")
wscript.echo mini.writeinivalue("Sec99","mivalor","3.141592")
mini.CloseINIFile()

Function ScriptPath()
    Dim path:  path = WScript.ScriptFullName
    ScriptPath = Left(path, InStrRev(path, "\"))
End Function


' ____________________ START INI Class HERE ________________________________________

Class ClsINI
'pierde los comentarios en el ini
'reordena los campos al salvar

Private FSO, TS, Dic, sFil,dirty 

 Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject")
 End Sub
          
  Private Sub Class_Terminate()
     Set FSO = Nothing
     Set Dic = Nothing
  End Sub
  
   '--Function to Read INI file into Dic: -------------------------------------
  Public Function OpenINIFile(sFilePath)
    Dim s, sSec, sList
     If FSO.FileExists(sFilePath) = False Then
       OpenINIFile = False
       Exit Function
     End If
     sFil = sFilePath
    Set Dic = Nothing  '-- reset Dic in Case an earlier file wasn't closed with CloseINIFile.
   
    Set Dic = CreateObject("Scripting.Dictionary")
     dim pref:pref= "[]" 
     dim equ
     On Error Resume Next 
     Set TS = FSO.OpenTextFile(sFil, 1)  
      Do While TS.AtEndOfStream = False
         s = Trim(TS.ReadLine)
         If Len(s) > 0 Then
            equ=instr(s,"=")
            If left(s,1)=";" then
              'comentario, saltamos linea
            elseIf Left(s, 1) = "[" Then
               pref=s
            ElseIf equ>1 Then
                 Dic.Item (pref&" "&left(s,equ-1))= trim(mid(s,equ+1))
            end if
         End If    
       Loop
      TS.Close  
     Set TS = Nothing
     OpenINIFile = True
  End Function
 
'------------------------------------------------------------------------- 
    
  Public Sub CloseINIFile()
     WriteNewINI()
     Set Dic = Nothing
  End Sub
     
'-------------------------------------------------------------------------

  'read one value from INI. return 0 on success. 1 If no such value. 2 If no such section.
  ' 3 If no file open. 4 If unexpected error in text of file.
Public Function GetINIValue(sSection, sKey)
    Dim s1: s1 = "["& sSection &"] "& sKey
    if Dic is Nothing then   GetINIValue = Null: Exit Function
      
    if not dic.exists(s1) then GetINIValue = Null:exit function 
    GetINIValue = Dic.Item(s1)
End Function



'--------- Write INI value: ---------------------------------
    ' return 0 on success. 2 If no such section.
    ' 3 If no file open. 4 If unexpected error in text of file. 
Public Function WriteINIValue(sSection, sKey, sValue)
    if Dic is Nothing then  WriteINIValue = 3: Exit Function
    Dim s1:s1 = "["& sSection &"] "& sKey
    Dic(s1)=sValue
    dirty=1	
    WriteINIValue=0
end function 
   
'---Function to delete single key=value pair: ---------------------------------------
   
Public Function DeleteINIValue(sSection, sKey)
    Dim s1:s1 = "["& sSection &"] "& sKey
    if Dic is Nothing then  DeleteINIValue = 3: Exit Function
    if not dic.exists(s1) then DeleteINIValue = 2:exit function 
    dic.remove(s1)
    DeleteINIValue =0
End Function
 
'-----------------------------------------------------------
   Private Sub WriteNewINI()  'ordenar y salvar
      const advarchar=200
      const adopenstatic=3
      Const fsoForWriting = 2
      dim i,s1,k1,n,rs,lastkey,sk
      if dirty=0 then exit sub 
      Set rs = CreateObject("ADODB.RECORDSET")
      with rs 
      
      .fields.append "SectionKey", adVarChar, 100
      
      .CursorType = adOpenStatic
      .open
      for each i in Dic.keys
       .AddNew
       rs("SectionKey").Value = i
       .Update
      next
      .Sort= " SectionKey ASC"
      
      .MoveFirst
      lastkey="[]"
      Set TS = FSO.OpenTextFile(sFil, fsoForWriting)  	
      do while not rs.EOF
        sk=rs("SectionKey")
        wscript.echo ">" & sk & "<"
        n=instr(sk,"]")
        s1=left(sk,n)
        k1=trim(mid(sk,n+1))   
        if s1<>lastkey  then
          lastkey=s1
          ts.writeline
          ts.writeline lastkey
        end if
        ts.writeline k1 & "=" & Dic(sk)
        .movenext
      loop
      .close
      end with
      set rs=Nothing						
      ts.close		      
      set ts=Nothing
    end sub
end class
 
' __________________ End INI Class HERE ______________________________