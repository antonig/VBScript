'Reads an IEEE764 floating point value from a string buffer
'
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
  if exponent=&h7F then   'nan o inf
     if mantissa=0 then 
        'bintofloat=sign * 1e38
        bintofloat="Inf"
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