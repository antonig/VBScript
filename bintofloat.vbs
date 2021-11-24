'retrieves an IEEE floating point number from an array of bytes. VBS can't do it by itself!!
function bintofloat (str,i)  'string, position
    dim n1,n2,exponent,mantissa,sign
    'only normalized, does'nt detect infinity or NaN
	const m_sign= &h8000, m_exp= &h7f80
	const d_exp= &h800000& , m_mant=&h7fffff&
	n1=asc(mid(str,i+3,1))* 256 + asc(mid(str,i+2,1))
         exponent = ((n1 and m_exp)/&h80)
	sign=1+ 2*((n1 and m_sign)<>0)
	n2 = ( asc(mid(str,i+2,1))* 256 + asc(mid(str,i+1,1)))* 256 + asc(mid(str,i,1))
	mantissa= (n2 and m_mant)
  wscript.echo mantissa,exponent
  if exponent=&hFF then   'nan o inf
     if mantissa=0 then 
        'bintofloat=sign * 1e38
        if sign then bintofloat="-Inf" else  bintofloat="Inf"
     else        
        bintofloat ="NAN "& mantissa
        'bintofloat=sign *1e38
     end if   
   elseif exponent =0 then 
    if mantissa=0 then
   	    bintofloat=0 
    Else  'denormalizad
      bintofloat=sign* mantissa * 2.^-149
    end if      
   else
     bintofloat =  sign* (mantissa or d_exp) * 2.^(exponent -150)
   end if   
end function
 
'testing code 
test=array( _
   array(45.32,"423547AE"), _
   array(-1,"BF800000"), _
   array(-123456,"C7F12000"),_ 
   array(-2,"C0000000"),_
   array(-1e38,"FE967699"),_
   array(-1e-38,"806CE3EE"),_
   array(0,"00000000"),_
   array("inf","7F800000"),_
   array("-inf","FF800000"),_
   array("NaN","FFC00000"),_
   array(1e-41,"00001BE0")_
   )
   
 'relenamos el buffer  
 buff=""
 for each a in test
  for i=7 to 1 step -2  '-little endian
    buff=buff & chr("&h"&mid(a(1),i,2))
  Next
 Next
wscript.echo len(buff)
pos=1
for each a in test
  wscript.echo bintofloat(buff,pos) & " Should be " & a(0)
  pos=pos+4
Next
 