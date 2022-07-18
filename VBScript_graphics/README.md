VBScript has no direct access to Windows Graphics interface.

However, by creating files and exiting to external viewers, graphics can be done programmatically, in a non-interactive way.
WYSIWIG is not possible in VBS, as far i know...

The samples:

rastergraphic.vbs 
 has a class that draws to a BMP (24 or 8 bit) file.
 The program supplies a palette for 8 bits but a custom one can be provided. 
 Only 4 primitives: pixel, line, circle and filled box at the moment
 The demo displays the palette, and draws on top of it a couple of lines and a circle

turtle_graphics.vbs: 
 has a class that issues SVG vectorial graphic commands to an HTML file
 It has a turtle graphics interface. The demo draws a Sierpinski triangle

A second option is to use the free vbs_gfx helper program that opens a 640x480 window and allows to set pixels in real time.
vbs_gfx was made by the french teacher Philippe Haubenestel in 2009-2011 and is available at http://tp.nexgate.ch/vbs_gfx/
The age of the program shows in the window size. Only two primitives: put pixel and draw line. Not supported anymore. 
And no mouse or keyboard input is returned to VBS... 

