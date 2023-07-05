'Lazy class implementing a persistent dictionary AGV2023

'the class has only:
' a constructor that creates the dictionary and if the filename passed exists it recovers the dict values from it
' a destructor that saves the dictionary to file before destroying it
'
' the dictionary is a public member so it can be accessed directly without the class implementing a wrapper for
'   each dictionary property or method


Option Explicit
Class pdict
  Dim d 
  private fn
  
  Public Default Function Init(filename)
    Dim dki,s,di,s1,dk,cnt
    Set d=CreateObject("scripting.dictionary")
    fn=Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\")) & filename
    WScript.Echo fn
    If CreateObject("Scripting.FileSystemObject").FileExists(fn) Then
      With CreateObject("ADODB.Stream")
        .mode=3 'r
        .open
        .type=2 'text
        .Charset="x-ansi"
        .lineseparator=10
        .loadfromfile fn
        s1=Split(.readtext ,Chr(10))   'no me deja hacer un readtext linea a linea y ponerla a un string
        cnt=1
        For Each s In s1
          debug.writeline s,Len(s),cnt
          cnt=cnt+1
          dki=Split(s,Chr(9))
          debug.WriteLine UBound(dki)
          If InStr(dki(1),Chr(1)) Then di=Split(dki(1),Chr(1)) Else di=dki(1)
          If Not d.Exists(dk) Then d.Add dki(0),di
        Next
      End With
    End If 
    Set Init=me
  End Function 

  Private Sub Class_Terminate() 
     Dim di
     with WScript.CreateObject("ADODB.Stream")
        .mode=3
        .Open
        .type=2    'texto
        .lineseparator=10
        .Charset = "x-ansi"
        For Each di in d.Keys
           If IsArray(d(di)) Then
            .writetext di & vbTab & Join(d(di),Chr(1)) ,1
           Else
            .writetext di & vbTab & d(di),1
           End If
        Next 
        .flush
        .position=.position-1  'matamos el ultimo eol para que al recuperar el split no nos de un item vacio de mas
        .seteos
        .SaveToFile fn ,2 'adSaveCreateOverWrite
     End With      
     Set d=Nothing
  End Sub
End Class

Sub print(s) 'prints to console arrays, dictionaries or simple variables. Adds a crlf at the end of the line
    Dim dk,s1 
    If IsArray(s) Then s=Join(s)
    If TypeName(s)="Dictionary" Then
       For Each dk in s.Keys 
          If isarray (s(dk)) Then s1=Join(s(dk),", ") Else s1= s(dk)
          WScript.stdout.WriteLine Join(Array(dk,s1),vbTab)
       Next 
    ElseIf IsArray(s) Then 
       WScript.stdout.WriteLine Join(s,", ") 
    Else  
      WScript.stdout.WriteLine (s)  
    End If
End Sub

'test code-----------------------------------------------
Dim pd,pdd

'si existe el archivo, rellena el dict con los valores en él
Set pd = (New pdict).Init("midict.txt")

'acceso directo al dict, si no la clase tendria que tener un envoltorio para cada propiedad o metodo
Set pdd=pd.d
If pdd.count=0 Then
   print "no se han recuperado valore para dict"
   pdd.add "hola","que pasa"
   pdd.add "valor",Array(1,2,3.15,4e12,5,6)
   
Else
  print "visualizando valores de archivo"   
End If
debug.WriteLine "aqui estamos"  
print(pdd) 
print pdd.count  'lle , solo pilla 1 valor
print pdd("hola")
print pdd("valor")
'exporta el diccionario a archivo
Set pd=Nothing


  