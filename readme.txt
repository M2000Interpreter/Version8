M2000 Interpreter and Environment

Version 9.8 Revision 28 active-X

Fix in syntax color (at edit) for parenthesis
look at info.gsb  module maze  line 37
maze$(INT((currentx% + oldx%) / 2), ((currenty% + oldy%) / 2)) = " "
The two last ) must have different color, but in previus revisions have the same. Now is ok

Syntax color also fixed for the same for EditBox (a control of M2000 GUI internal system)


The fist time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory




From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate


https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

https://www.dropbox.com/s/xt30bspw6q9pf5f/M2000language.exe?dl=0

http://georgekarras.blogspot.gr/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             