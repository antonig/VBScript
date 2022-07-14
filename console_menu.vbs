'simple foolproof menu for console

'define the keys using constants
const optionA="A"
const optionB="B"
const optionC="C"

'the menu text built using constants. 
'the .- after the key letter will allow us to validate input
mm="Menu"& vbcr & _ 
optionA &".- option A"& vbcrlf & _ 
optionB &".- option B"& vbcrlf & _ 
optionC &".- option C"& vbcrlf 

wscript.echo mm

'get key and validate it. (Enter is a hidden valid option)
do
  wscript.stdout.write "Select an option and [Enter]"
  s=  wscript.stdin.readline 
  s= ucase(s)     'remove if case sensitive
loop until instr(mm,s & ".-")>0   'validate "[key].-" appears in the menu text

'the switchboard, using the constants
select case s
case optionA  
  wscript.echo "You selected Option A"
case optionB  
  wscript.echo "Option B"
case optionC  
  wscript.echo "Option C"  
case ""                      'using the [Enter] hidden option as an abort            '
  wscript.echo "quitting"
  wscript.quit(0)
case else
  wscript.echo "you should not have reached this"
end select
