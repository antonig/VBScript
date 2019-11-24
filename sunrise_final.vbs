option explicit
'-------------------------------------------------------------------------------------
' Calculo hora salida y puesta de sol 
' 2017 Controlli Iberica AGV 20/12/2017 V1.1
' 
' USO de la función: 
'  x = aman_anoch(dateserial(año,mes,dia),latitud, longitud,ang1,ang2) 
'
'     dateserial convierte la fecha a formato interno VBS
'     latitud y longitud  en grados con decimales (E y N positivos, W y S negativos) pueden copiarse valores de la URL en Google Maps
'     la latitud debe ser en rango +/-66.25 (zona en que el sol sale y se pone cada dia - se excluyen zonas polares)
'     ang1,ang2   grados de correccion  ang1: al amanecer ang2: al anochecer 
'        0  da al hora a la que asoma /se esconde el sol
'        6  da la hora del crepusculo civil   ( se ve lo suficiente como para realizar actividades )
'        12 da la hora del crepusculo nautico ( esta oscuro pero  se ve el horizonte )  
'        +/-correccion adicional por altura del horizonte  (montañas)  
'            negativo si horizonte esta mas alto que observador, positivo si esta mas bajo
'
'  la función devuelve un array  de 4 valores de tiempo corregidos para hora civil y hora de verano
'       en x(0) hora civil  de amanecer  ( en timeserial VBS )  
'       en x(1) hora civil  de anochecer ( en timeserial VBS ) 
'       en x(2) hora civil de mediodia   ( en timeserial VBS ) 
'       en x(3) longitud del dia         ( en timeserial VBS ) 
'
'NOTA:  la función clava  (+/- 1 min)  los datos de las tablas generadas por http://aa.usno.navy.mil/data/docs/RS_OneYear.php
'NOTA2: si usamos WMI para saber si hoy estamos en DST solo podemos calcular el dia de hoy  
'           adjunto funciones para calcular DST para cualqueir sitio de Europa
'-------------------------------------------------------------------------------------
const DosPi=6.283185307179586476925286766559
dim UsarMiDST,mi_GMToffset  'se usa para seleccionar expresion en la funcion Solar_ACivil

'-------------------------------------------------------------------
' calculos hora astronomica a hora civil (husos horarios, DST)
'-------------------------------------------------------------------

' DST opcion A : obtener estado DST de PC: solo para dia de hoy
function GMToffsetDST() 'detos de hora civil obtenidos de Windows mediante WMI
  ' devuelve:  r(0) offset huso horario en minutos y r(1) si estamos en DST 
  dim objWMIService,colItems,objItem,colItems2,objItem2,GMToffset,enDST, r(1)
  Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
  Set colItems = objWMIService.ExecQuery("Select * From Win32_TimeZone")
  For Each objItem in colItems
    r(0)=objItem.Bias 
    Set colItems2 = objWMIService.ExecQuery("Select * From Win32_ComputerSystem")
    For Each objItem2 in colItems2
        r(1) = objItem2.DaylightInEffect
    Next
  Next
  GMToffsetDST=r   
end function  

'DST opcion B:  para poder hacer calculos a largo plazo (no solo para hoy)
'UsarMiDST=1

function Calc_dias_DST(Anno) ' obtenemos el dia ¡¡¡EUROPEO!!! de DST para este año(ultimo domingo de marzo y de octubre)
	dim Startdate,EndDate,dias(1)
	'para europa, ultimo domingo de marzo y de octubre 
	StartDate = DateSerial(Anno, 3, 31)
	EndDate = DateSerial(Anno, 10, 31)
	dias(0)=StartDate -( WeekDay(StartDate) - vbSunday )+ timeserial(2,0,0)
	dias(1)=EndDate - (WeekDay(EndDate) - vbSunday)+ timeserial(2,0,0) 
	Calc_dias_DST=dias
end function

Function FenDST(eldia) 
  dim dst_dias 
  DST_dias = Calc_dias_DST(Year(eldia))
	FenDST = (eldia > DST_Dias(0) and  eldia < DST_Dias(1) )
End Function

'
function solarAcivil(date, lon) 'diferencia entre hora astronomica y hora civil  en timeserial
  'Salida: x(0)=GMToffset en minutos,, x(1)= 1 si estamos en DST 
  dim x  
  If isempty(UsarMiDST) then
    x= GMToffsetDST()                                         'usar esta linea en produccion (toma datos de windows)
  else
    redim x(1)
    x(0)=Mi_GMToffset: x(1)= FenDST(date+timeserial(12,0,0))  'esta es para pruebas usa funciones Calc_dias_DST y FenDST 
  end if  
  solarAcivil=timeserial(0,x(0)-lon*4 - (x(1)*60),0)
end function

'----------------------------------------------------------------------
'calculos astronómicos
'--------------------------------------------------------------------



function doy2rad(doy)
   doy2rad=doy*DosPi/365
end function   

function rad2deg(rad)
  rad2deg=rad*360/DosPi
end function

function deg2rad(deg)
  deg2rad=deg*DosPi/360 
end function

function arccos (x)  ' arco coseno en radianes
  IF x = 0 THEN arccos = DosPi/4: EXIT FUNCTION
  IF x > 0 THEN
    arccos = ATN(SQR(1 - x * x) / x)
  ELSE
    arccos = DosPi/2 + ATN(SQR(1 - x * x) / x)
  END IF	
end function 

function doy(date) 'dias desde el 1 de enero
  doy=date-dateserial(year(date),1,1)+1
end function  
 
function eot(dy) 'ecuacion de tiempo (desplazamiento de mediodia solar respecto mediodia teorico en timeserial VBS
  'varia +/- 14 min durante año 21-3 -7m 5s | 21-6  -1m 51s  |21-9  6m59s |21-12  1:49
  dim b,c
  b=doy2rad(dy)
  c=229.18*(0.00075+0.001868*cos(b)-0.032077*sin(b)-0.014615*cos(2*b)-0.040849*sin(2*b))
  eot=timeserial(0,int(c),60*(c-int(c)))
end function

function decli(dy) 'declinacion del sol en radianes
  '21/3 y 21/9 es 0, 21/9 es +23.45,  21/12 es -23.45
  dim c
  c=doy2rad(dy)
  decli = 0.006918 - 0.399912*cos(c)+ 0.070257*sin(c) - 0.006758*cos(2*c)+ 0.000907*sin(2*c) - 0.002697*cos(3*c)+ 0.00148*sin(3*c)
  'decli=deg2rad(23.75)*sin(doy2rad(dy+284))
end function  

function dlength(dy,lati,crepu) ' duracion de medio dia  en timeserial  
  'formula de wikipedia
  dim c,d,h
  d=decli(dy)
  h=sin(deg2rad(-0.83- crepu))   ' 
  c=4*rad2deg(arccos((h-sin(d)*sin(deg2rad(lati)))/cos(d)/cos(deg2rad(lati))))
  dlength=timeserial(0,int(c),60*(c-int(c)))
end function


function aman_anoch(date,lati, longi,a_rise,a_set)  'devuelve array 2 valores con hora amanecer y hora anochecer en timeserial
  
  dim mi_doy,mi_largo1,mi_largo2,mi_eot,mi_offset,mi_decli,p(3),mi_mediodia,s
  if abs(lati)> (90-23.75) then
     s=  " latitud fuera del rango < +/- 66.25" 
     msgbox s,0,"Error" 
     p(0)=s    
  else  
    mi_doy   = doy(date)   'en radianes
    mi_eot   =  eot(mi_doy)
    mi_largo1 =  dlength(mi_doy,lati,a_rise) 
    mi_largo2 =  dlength(mi_doy,lati,a_set) 
    mi_offset= solarACivil(date,longi)  'minutos
    mi_mediodia=timeserial (12,0,0)-mi_eot +mi_offset 'datetime 
    'msgbox(date & vbcrlf & cdate(2*mi_largo) & vbcrlf & mi_offset & vbcrlf & mi_eot & vbcrlf & mi_mediodia & vbcrlf & rad2deg(decli(mi_doy)))  
    p(0)   = cdate(mi_mediodia-mi_largo1) 
    p(1)   = mi_mediodia+mi_largo2 
    p(2)   = mi_mediodia
    p(3)   = mi_largo1+mi_largo2
  end if 
  aman_anoch=p
end function 

'ejemplo de uso
dim x
x=aman_anoch (int(Now),41.39,2.16,0,0) 'prueba con latitud y longitud de Barcelona
msgbox  int(now) & vbcrlf & "Amanece a las " &  x(0) & vbcrlf & "Anochece a las " & x(1) & vbcrlf & vbcrlf & _
"Mediodia a las " & x(2) & vbcrlf & "Duracion del dia " & x(3) ,0,"Valores para hoy en Barcelona"

'fin del codigo astronómico
'------------------------------------------------------------------------------------------------

'--lo que viene a continuacion son pruebas , puede eliminarse ------------------------------------------------
dim d_astro, crepusculo

sub prueba(lat,longi,ang1,ang2) 'calculamos para solsticios y equinoccios y visualizamos errores respecto tabla 
  dim s,m, temp
  s="Amanecer" & vbtab &"Err" & vbtab & "Anochecer" & vbtab &"Err " & vbcrlf
  dim cnt:cnt=0
  for m=3 to 12 step 3  'probamos tres dias del año, equinoccios y solsticios
    dim dt : dt=dateserial(2017,m,21)
    temp=aman_anoch(dt,lat,longi,ang1,ang2)
    s=s & dt & vbcrlf &  temp(0)& vbtab & cint(1440*(temp(0)-d_astro(cnt))) & vbTab & temp(1)& vbtab & cint(1440*(temp(1)-d_astro(cnt+1)))  & vbcrlf & temp(2) & vbtab & temp(3) & vbcrlf
    cnt=cnt+2
  next
  wscript.echo(s)
end sub


sub pruebaciudades ( )  'probamos funcion respecto a tablas para varias ciudades  http://aa.usno.navy.mil/data/docs/RS_OneYear.php
  dim latitud,longitud,alto,c
  c= inputbox(" 0 Barcelona "& vbcrlf & " 1 Santa cruz de tenerife" & vbcrlf & " 2 Hamburgo" & vbcrlf & " 3 Buenos Aires ","Seleccionar ciudad ",0)  

  select case c
    'datos de Barcelona
  case 0:
    d_astro=array( _
      timeserial(6,53,0),timeserial(19,05,0),_
      timeserial(6,18,0),timeserial(21,28,0),_
      timeserial(7,38,0),timeserial(19,50,0),_
      timeserial(8,14,0),timeserial(17,25,0))
    Mi_GMToffset=60
    latitud = 41.39     '41ª23'  Norte
    longitud =2.16     '2º11'  Este 
  case 1:
    'datos de Santa Cruz de Tenerife 
    d_astro=array( _
      timeserial(7,08,0),timeserial(19,17,0),_
      timeserial(7,08,0),timeserial(21,05,0),_
      timeserial(7,53,0),timeserial(20,03,0),_
      timeserial(7,53,0),timeserial(18,13,0))
    Mi_GMToffset=0
    latitud=28.46
    longitud=-16.25
  case 2:
    'datos de Hamburgo 53.5477802,10.0111023
    d_astro=array( _
      timeserial(6,20,0),timeserial(18,35,0),_
      timeserial(4,50,0),timeserial(21,53,0),_
      timeserial(7,04,0),timeserial(19,21,0),_
      timeserial(8,34,0),timeserial(16,02,0))
    Mi_GMToffset=60
    latitud=53.55
    longitud=10.01
  case 3:
    'datos de Buenos Aires -58.3815591 W -34.6036844S
    d_astro=array( _
      timeserial(6,57,0),timeserial(19,03,0),_
      timeserial(9,00,0),timeserial(18,50,0),_
      timeserial(7,44,0),timeserial(19,49,0),_
      timeserial(5,37,0),timeserial(20,06,0))
    Mi_GMToffset=-180
    latitud=-34.60
    longitud=-58.38
  case else:
    msgbox "ninguna ciudad seleccionada"
    exit sub   
  end select
  prueba latitud,longitud,0,0
end sub

pruebaciudades  
