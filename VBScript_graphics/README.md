# VbScript Graphics

VBScript has no direct access to the Windows Graphics interface. To get graphics there are some workarounds

## The first option is
To  create a graphics file and exit to its externa viewer.This way graphics can be done programmatically, in a non-interactive way.
WYSIWIG is not possible in VBS, as far i know...

The samples cover two different approaches:

### rastergraphic.vbs 
 has a class that draws to a BMP (24 or 8 bit) file.
 The program supplies a palette for 8 bits but a custom one can be provided. 
 Only 4 primitives: pixel, line, circle and filled box at the moment
 The demo displays the palette, and draws on top of it a couple of lines and a circle

### turtle_graphics.vbs: 
 has a class that issues SVG vectorial graphic commands to an HTML file.
 It has a turtle graphics interface. The demo draws a Sierpinski triangle

## A second option is to use the free vbs_gfx helper program 
It opens a 640x480 window and allows to set pixels  in real time.
vbs_gfx was made by the french teacher Philippe Haubenestel in 2009-2011 and is available at http://tp.nexgate.ch/vbs_gfx/
The age of the program shows in the window size. Only two primitives: put pixel and draw line. It's not supported anymore. 
And no mouse or keyboard input is returned to VBS... 

## The third option is to use textmode with ANSI escape codes
 This allows to locate text, to get colors or box drawing characters. 
 Unfortunately windows from 2000 to 8.1 don't recognise ANSI codes, that show as garbage on the console.

### conways_life.vbs 
is a sample of an app using ANSI graphics dynamically.



