# vbs
Some subs and functions I find useful. And my try on doing graphics with VBS 


## bintofloat
Gets  4bytes from a binary stream in a string and reads them as an IEEE764 floating point number

## clean_thml
Removes nontext tags of an html file. UTF-8

## console_menu    
A "foolproof" console menu. The accepted inputs are defined as constants. The menu text and the switchboard use these   constants.         The input validation find the characters in the menu text yhet's built from the constants.


## general_lib      
Assorted general use routines. 
  Loads text file ascii or utf-8. Splits it to array if required. 
  Get string or array (joins it to string) and saves to file ascii or utf-8 and open it in notepad
  Select file. reads name from first argument. If not exists open windows file selector
  Ensure running in cscript/wscript or 32/64 bits version of vbscript. Restarts script if required.
  Read a recordset from  csv or from xls excel file 
  Display selected columns of given width from recordset. Right-left align each column at will
  Check if service is running


## lista_recordset 
uses odbc to read a csv to a disconnected recordset, then uses an array to tell a sub how to print it in columns.

## mi_ClassIni      
An exercise in using classes. It reads an ini file to a dictionnary making items of the ini available to program. The progam can add, modify and delete items. At the end the class allows to save back the values to the ini. Warning: If the ini file is written back it loses the comment lines!
                  
## my_userinput  
function that gets an array defining an input form with text /radio buttons/checkboxes, builds an html form and opens it with InternetExplorer.Aplication. On submit it reads the fields to a dictionnary. Warning: Esthetic performance may vary.  No validation is performed while form is open.
                  
## oneliners.txt 
Some snippets i find useful               

## prio
Handy priority queue class.

## serialize_dict
Class To save And recover a dictionary To/from a file. The dictionary can have Array items

## srt_time_offset
Adds/substracts an offset to all times in a SRT video subtitle file.

## sunrise_final
Calculates sunrise and sunset time in civil time or any position under the polar cicles. Has an angle compensation for if the place is not at the same height as the visible horizon (mountins, top of a tower)
Gives results +/- 1 min of the results provides by the US Naval Observatory site.

## printf sprintf
A try to get straighforward string formatting functions copying c#'s string interpolation

## rosetta_notdone
Web scrapping Rosetta code to list tasks not done in VBScript

## misudoku2
Sudoku solver using simple strategies and if needed, brute force. Can get problems from text files in different formats. See command line options.

