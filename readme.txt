M2000 Interpreter and Environment

Version 9.9 revision 11 active-X
Better code on passing by reference static variables and array items,
now return error if these passing aren't in a module or function call, also for
subroutines excluded these, now raising error.

The Mutex object now works fine. Clock module in info has an example for it.
The clock can't run two times, because the second one check the mutex and exit early.

George Karras, Kallithea Attikis, Greece.

The fist time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory


To use Help (an mdb database) you need Access 2007 runtime.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)


https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

https://www.dropbox.com/s/xt30bspw6q9pf5f/M2000language.exe?dl=0

http://georgekarras.blogspot.gr/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             