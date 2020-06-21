M2000 Interpreter and Environment

Version 9.9 revision 31 active-X
1. New With operator for combining groups is right expressions
(combining at object level)
Class alfa {x=10}
Class beta {y=30}
Z=alfa() with beta()
List ' show Z.X and Z.Y

2. Combining classes from classes (at definition level)
Class A {X=10}
Class B {Y=30}
Class AB as A as B {Z=100}
M=AB()
List  ' show M has three members M.X, M.Y, M.Z

3. Reduce copy on group assignment (without the Set member - which control assignment - the Valid(Z.Z) return True)
Class A {
	X=10
	Set {read This} ' this reduce the copy
}
Class B as A {Z=10}
M=B()
M.X=500
Z=A()
Z=M
Print Z.X=500, Valid(Z.Z)=False


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

                                                             