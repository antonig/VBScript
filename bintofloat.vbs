'retrieves an IEEE floating point number from an array of bytes. VBS can't do it by itself!!
function bintofloat (str,i)  'string, position
    dim n1,n2,exponent,mantissa,sign
    'only normalized, does'nt detect infinity or NaN
	const m_sign= &h8000, m_exp= &h7f80
	const d_exp= &h800000& , m_mant=&h7fffff&
	n1=asc(mid(str,i+3,1))* 256 + asc(mid(str,i+2,1))
	if n1  and m_sign then sign=-1 else sign=1
	n2 = ( asc(mid(str,i+2,1))* 256 + asc(mid(str,i+1,1)))* 256 + asc(mid(str,i,1))
	exponent = ((n1 and m_exp)/&h80)
	mantissa= (n2 and m_mant)
  if (exponent and &hFF)= &hFF then 'nan
     bintofloat ="NAN "& mantissa
	elseif (exponent or mantissa) =0 then 
	   bintofloat=0 
	else
     bintofloat =  sign* (mantissa or d_exp) * 2.^(exponent -150)
	end if   
end function
 