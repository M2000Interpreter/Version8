M2000 Interpreter and Environment

Version 9.9 revision 1 active-X

1. update the socket and the client objects. See down4 example in info.gsb.
2. New read only variable Internet$ return the public ip or 127.0.0.1 if no internet connection exist
3. New read only variable Internet return true if there is an internet connection
4. New #Eval() and #Eval$() for tuples.
5. Report statement rewrite for use proper the TAB character when use justification. Also now is faster. Use a variant of code from internal editor to find the proper point to cut the line, using a binary search. The old wwplain function exist as wwplainOLD, for information only.
6.Improved Test form (a bug fixed).
7. Many improvments and bug fixed. 
8. Many examples in info.gsb also improved, and new examples added.

George Karras, Kallithea Attikis, Greece.

The fist time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory




From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)


https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

https://www.dropbox.com/s/xt30bspw6q9pf5f/M2000language.exe?dl=0

http://georgekarras.blogspot.gr/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             