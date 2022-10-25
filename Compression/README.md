# Compression
File compression in VBS

## lzw.vbs
An implemantation of the LZW algorithm. It uses 2 bytes per code, variable length codes are out of question in VBS, as those would require lots of integer division and modulus operations. Not so fast, compresion requires 5 to 10 segs per MB. Decompression is much faster. 

## ascii-art-rle.vbs
Uses a homebrew variant of the RLE line algorithm to compress ASCII art images with a lot of spaces. Not so good with images with shade gradients. It encodes pure ascii unique chars as 128+asc and lengths with the same char as 128+asc & 32+length.


