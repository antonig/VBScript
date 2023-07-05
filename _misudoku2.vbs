Option Explicit
'VBScript Sudoku solver. AGV202301
'Uses first single candidate strategies then a backtrace recursive algorithm if required. 
'can read a problem from: 
'  if no parameters passed or an error loading external prolem occurs, it solves a built in problem  (see string prob0).
'  a problem can be included in the command line
'  a file can be read with the argument /f:textfile
'  if file has a sudoku in each line, use /l:line to select the line to read (don't use if a problem uses more than one line)
'  ****use /noansi: in the command line if you screen displays garbage  (Windows version < Win10 or running in an IDE)***
'A problem string can have 0's or dots in the place of unknown values. All chars different from .0123456789 are ignored
'  newlines are allowed in the definition except if the /l:line parameter is used

'TODO: 
'ok  visualizar valores iniciales en resultado
'ok  usar estrategias simples antes de atacar con backtrace
'    calcular 27 mapas de cifras usadas al principio y mantener actualizado en recursion en lugar de calcular n veces
'    detectar cond iniciales imposibles 
'    implementar eliminacion de linea y de bloque 

'display
Dim g_noansi :g_noansi=True
Const ansblu="[94m"
Const ansblk="[0m"
Dim grid
Const gr="+a+a+a+"
grid=array(Replace(gr,"a",String(7,"-")),Replace(gr,"a",String(31,"-")),Replace(gr,"a",String(7,"-")))

'bit masks
Const c_all=&h3FE  'masks all candidates
Dim pwr:pwr=Array(1,2,4,8,16,32,64,128,256,512,1024,2048)

'the default problem
Dim prob0:prob0= "001005070"& "920600000"& "008000600"&"090020401"& "000000000" & "304080090" & "007000300" & "000007069" &  "010800700"

'work arrays
Dim s0(8,8),sdku(8,8),cand(8,8),opt(8,8)

'general use routines
Sub pause
  wscript.stdout.write "Enter to continue: ": wscript.stdin.read 1
End Sub

Sub print(s): 
  On Error Resume Next
  WScript.stdout.Writeline (s)  
  If  err= &h80070006& Then WScript.Echo " Please run this script with CScript": WScript.quit
End Sub

Function pad(s,n) 
  If n<0 Then pad= right(space(-n) & s ,-n) Else  pad= left(s& space(n),n)  
End Function

Function bitcount16(ByVal c)  'VALIDO HASTA 16 BIT 
  c= C-((c\2) And &H5555&)
  C=((C\4) And &H3333&)+(C And &H3333&)
  C=((C\16)+C) And &h0F0F&
  bitcount16 =((C\256)+C) And &HFF&
End Function

'reading problem
Function parseprob(s)'problem string to array
  'returns number of clues in problem,0 if string does'nt have 81 positions
  Dim i,j,m,cnt,row,col
  print "parsing: "&vbcrlf & s & vbCrLf
  resetgrid
  
  j=0
  For i=1 To Len(s)
    col=j mod 9
    row=j \ 9
    m=Mid(s,i,1) 
    Select Case m
      Case "1","2","3","4","5","6","7","8","9"
      s0(row,col)=CInt(m)
      modif row,col,CInt(m),True
      j=j+1:cnt=cnt+1
      Case ".","0"
      j=j+1
      Case Else  'all other chars are ignored as separators
    End Select
  Next
  
  If j<>81 Then 
    parseprob=0
    print "The problem entered does'nt have data for a 9x9 sudoku. Got " & j & "cases. Solving default problem"
  Else 
    parseprob=cnt
  End If  
End Function      

'command line
Function getprob  'get problem from file or from command line or from the string prob0
  'returns number of clues in problem
  Dim s,s1,n,i
  With WScript.Arguments.Named
    On Error Resume Next
    If .Exists("noansi") Then g_noansi=True
    If .exists("f") Then
      s1=.item("f")
      If InStr(s1,"\")=0 Then s1= Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))&s1
      If .exists("l") Then
        n=.item("l") 
        print "reading problem in line " & n &" from file "&s1
        err.clear
        With CreateObject("Scripting.FileSystemObject").OpenTextFile (s1, 1)
          For i=1 To n-1 : 
            .skipline
          Next
          If err Then 
            print "The requested line " & n &" does'nt exist in file " & s1 & " Solving default problem"  :  
            getprob=parseprob(prob0): 
            Exit Function
          End If  
          s=.readline  
          getprob=parseprob(s1) :If getprob<>0 Then Exit Function
        End With
      Else
        print "reading problem from file "&s1
        s= CreateObject("Scripting.FileSystemObject").OpenTextFile (s1, 1).readall
        If err Then  print "can't open file. " & s1 & " Solving default problem"  :  getprob=parseprob(prob0): Exit Function
        getprob=parseprob(s):If getprob  Then Exit Function
      End If 
    End If   
  End With
  With WScript.Arguments.Unnamed            
    If .count<>0 Then
      print "reading problem from command line"
      s1=.Item(0)
      getprob=parseprob(s1) :If getprob<>0 Then Exit Function
    End If  
  End With   
  getprob=parseprob(prob0)
End Function

Function first(x,n) 
  first=n-1 
  Do: 
    first=first+1: 
  Loop until x And pwr(first):
End Function 

Function decode(x) 
  Dim i
  decode=""
  For i=1 To 9
    If (pwr(i) And x) Then decode=decode & i 
  Next  
End Function

Sub resetgrid  'first scan check for candidates
  Dim r,c
  For r=0 To 8
    For c=0 To 8
      sdku(r,c)=0
      cand(r,c)=c_all
      opt(r,c)=9
    Next  
  Next
End Sub

Sub modif (row,col,cod,add) 
  'añade o quita un valor cod en posicion pos, actualizando mascaras y contajes de fila,col y bloque 
  Dim n,r,c, blkr,blkc,inc
  blkr=(row\3)*3
  blkc=(col\3)*3
  If Not add Then
    n=pwr(cod)
    inc=+1
    sdku(row,col)=0
    
  Else 'add
    n=Not pwr(cod) And c_all
    inc=-1
    sdku(row,col)=cod
    cand(row,col)=0
    opt(row,col)=0
  End If 
  For r=0 To 8
    If r<>row Then
      If add Then  cand(r,col)=cand(r,col) And n Else cand(r,col)=cand(r,col) Or n
      If sdku(r,col)=0 Then opt(r,col)=opt(r,col)+inc
    End If
  Next  
  For c=0 To 8   
    If c<>col Then
      If add Then  cand(row,c)=cand(row,c) And n Else cand(row,c)=cand(row,c) Or n
      If sdku(row,c)=0 Then opt(row,c)=opt(row,c)+inc
    End If
  Next
  For r=blkr To blkr+2
    For c=blkc To blkc+2
      If (r <> row) And (c<>col) Then
        If add  Then cand(r,c)=cand(r,c) And n Else cand(r,c)=cand(r,c) Or n
        If sdku(r,c)=0 Then opt(r,c)=opt(r,c)+inc
      End If 
    Next  
  Next
  For r=0 To 8
    For c=0 To 8
      opt(r,c)= bitcount16( cand(r,c))
    Next
  Next
End Sub


Function solve1  'un solo candidato, ok
  Dim r,c,cnt
  For r=0 To 8
    For c=0 To 8
      If opt(r,c)=1 Then  modif r,c,first(cand(r,c),1),True:cnt=cnt+1
    Next
  Next
  solve1=cnt        
End Function

Function solveunic  
  Dim j,r,c,cnt,rblk,cblk,b,x
 'filas
  For r= 0 To 8 
    ReDim a(9)
    For c=0 To 8
      x=cand(r,c)
      For j=1 To 9
        If x And pwr(j)  Then
          If IsEmpty(a(j)) Then 
            
            a(j)=Array(r,c)
          ElseIf IsArray(a(j)) Then  
            a(j)=1
          End If  
        End If
      Next
    Next
    For j=1 To 9
      'If cnt1=8 Then If IsEmpty(a(j)) Then modif a(j)(0),a(j)(1),j,True:cnt=cnt+1 
      If IsArray(a(j)) Then  modif a(j)(0),a(j)(1),j,True:cnt=cnt+1 
    Next
  Next
  'columnas
  
  For c= 0 To 8 
    ReDim a(9)
    For r=0 To 8
      x=cand(r,c)
      For j=1 To 9
        If x And pwr(j)  Then
          
          If IsEmpty(a(j)) Then 
            a(j)=Array(r,c)
          ElseIf IsArray(a(j)) Then  
            a(j)=1
          End If 
        End If
      Next
    Next
    For j=1 To 9
      If IsArray(a(j)) Then  modif a(j)(0),a(j)(1),j,True:cnt=cnt+1 
    Next
  Next
  'bloques
  
  For b=0 To 8
    ReDim a(9)
    rblk= (b\3)*3
    cblk=(b mod 3) *3
    For r= rblk To rblk+2
      For c=cblk To cblk+2
        x=cand(r,c)
        For j=1 To 9
          If x And pwr(j)  Then
            If IsEmpty(a(j)) Then 
              a(j)=Array(r,c)
            ElseIf IsArray(a(j)) Then  
              a(j)=1
            End If 
          End If
        Next
      Next
    Next   
    For j=1 To 9
      If IsArray(a(j)) Then  modif a(j)(0),a(j)(1),j,True:cnt=cnt+1 
    Next
  Next
  solveunic=cnt    
End Function

'backtrace solver
Function solve(x,ByVal pos)
  Dim row,col,r,c,used,i,r1,c1
  solve=False
  If pos=81 Then solve= True :Exit Function
  row= pos\9
  col=pos mod 9
  If x(row,col) Then solve=solve(x,pos+1):Exit Function
  
  cnt=cnt+1
  used=0
  For i=0 To 8
    used=used Or pwr(x(i , col))
    used=used Or pwr(x(row , i))
  Next
  r1 = (row\ 3) * 3
  c1 = (col \3) * 3
  For r=r1 To r1+2
    For c=c1 To c1+2 
      used = used Or pwr(x(r,c))
    Next
  Next
  For i=1 To 9
    If (used And pwr(i))=0 Then 
      x(row,col)=i
      solve= solve(x,pos+1) 
      If solve=True Then Exit Function
    End If   
  Next
  x(row,col)=0
  solve=False
  'WScript.StdOut.Write pos & " "
End Function

'display 0 displays the sudoku 
'display 1 displays array of candidates, 
'display 2 shows array of candidate counts
Sub display(n)
  Dim r,c,s
  For r=0 To 8
    If (r mod 3)=0 Then print grid(n)
    s=""
    For c=0 To 8
      
      If (c mod 3)=0 Then s=s & "| "
      Select Case n
        Case 0
        If sdku(r,c) Then
          If g_noansi= False Then If s0(r,c) Then s=s&ansblk Else s=s&ansblu
          
          s=s& pad(sdku(r,c),2)
          If g_noansi= False Then  s=s&ansblk
        Else
          s=s& "  "
        End If  
        Case 1 
        s=s& pad(decode(cand(r,c)),-10)
        Case 2 
        s=s& pad(opt(r,c),-2)
      End Select 
    Next
    print s&"|"
  Next
  print  grid(n)
  If g_noansi= False Then print ansblk
End Sub

Function pdtesolucion
  Dim r,c,cnt
  cnt=0
  For r=0 To 8
    For c=0 To 8
      If sdku(r,c)=0 Then cnt=cnt+1
    Next
  Next
  pdtesolucion=cnt
End Function    

Dim Time,cnt,cnti ,c1,c2
cnti=getprob
print cnti
print "The problem"  
display(0)
'display(1)
'display(2)

Time=Timer
print "Solving"
Do 
  c1= solveunic
  c2= solve1
Loop until (c1+c2)=0
If pdtesolucion Then solve sdku,0 :print ""

display(0)
If pdtesolucion Then
  display(1)
  display(2)
End If

print vbcrlf &"pending cases "& pdtesolucion & " from " & 81-cnti & " time: " & Timer-Time & " seconds "
pause