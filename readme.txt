M2000 Interpreter and Environment

Version 9.9 revision 47 active-X
1. Fix the scroll bar action in center mode in EditBox
2. Added logic to fix this problem by example:
Module Z (X) {Print "ok", X}
Z=10  ' now we have a module and a variable with same name
Z -10  ' THIS NOW EXECUTED AS MODULE CALL TO Z
before this revision the - operator can't used by variable Z so we got an error
now before the error raised interprer look if it is a module with same name
and if find then call it else raise the error.

George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com


The fist time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)


https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

https://www.dropbox.com/s/xt30bspw6q9pf5f/M2000language.exe?dl=0

http://georgekarras.blogspot.gr/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             