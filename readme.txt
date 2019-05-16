M2000 Interpreter and Environment

Version 9.8 Revision 20 active-X


1. New function #slice()
Print (1,2,3,4)#slice(0,1)  ' 1 2
Print (1,2,3,4)#slice(0,1)#rev() ' 2 1

2. a=(1, a$>"a", 2) now works (no problem with string comparison)
a$="b"
a=(1, a$>"a", 2)
print a  ' 1 True 2


3. correction in syntax color procedure for multiline strings in comparisons



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

                                                             