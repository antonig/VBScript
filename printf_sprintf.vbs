'printf y sprintf formato string interpolation de c#
' to do
'   añadir poner padding
'   letra de formato Hex Octal Dec F
'   \\n  deberia dar \n y da vbcrlf

function pad(s,n) if n<0 then pad= right(space(-n) & s ,-n) else  pad= left(s& space(n),n) end if :end function

function formats(byval s)
'only \\ \n \t
'only {expression} no  :padding or :format letter
	with new regexp
		.pattern="[\\]" 
		if .test(s) then
			s=replace (s,"\\","\") ' 
			s=replace(s,"\n",""" & vbcrlf & """)
			s=replace(s,"\t",""" & vbtab & """)
		end if	
		.pattern="[{}]"
		if .test(s) then
			s=replace(s,"{{","{")
			s=replace (s,"}}","}")
			.pattern="\{(.+?)\}"
			.global=true
			while .test(s)
				s=.replace(s,""" & $1 & """)
			wend
		end if 
	end with
  formats=""""&s&""""
end function

function sprintf(byval s)  sprintf=eval(formats(s)) :end function

sub printf(byval s)   wscript.stdout.write eval(formats(s)) :end sub	


a1=12
a2="pepe"
wscript.echo sprintf("hola {a1} \n que pasa {a2}")

printf("{a2} no pasa nada\\nada\t pero son las {a1}\n y ya es tarde")