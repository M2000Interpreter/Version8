M2000 Interpreter and Environment

Version 9.9 revision 33 active-X
1. Classes have types  (classes are functions which produce groups)
2. Groups may have type (more than one)
3. Group private variables can be used as public in a group method of a group of same types
4. Private variables can be passed by reference
5. we can define type in a parameter list:
	abc as *alfa means abc has to be pointer to alfa type
	abc as alfa means abc has to be a group of alfa type
6. Pointer() return a pointer to the Null type Group.
	Group(Pointer()) return the Null type Group
7 New operator "is type"
	abc is type alfa
	return true if one of types of abc is alfa

There are some new modules in info file using the new additions


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com


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

                                                             