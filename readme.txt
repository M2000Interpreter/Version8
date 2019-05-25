M2000 Interpreter and Environment

Version 9.8 Revision 24 active-X


Some optimizations.
Clear statement work better now, inside block.
We can't clear variables inside a block for temporary definitions (like in Subs and if "For This" structure)
Exit For break blocks until exit for structure.
Exit for get an optional label to perform a goto.

From inside If Then End if we can jump inside it, not outside. Using If Then {  } we can jump anywhere.

\\ jump anywhere (inside a module or function)
\\ using if then {} and for { } structures
if true then {
	for i=1 to 10 {
		if i>5 then 50
		print i
	}
	print "no this"
}
050	Print "ok"
\\ jump inside if then/end if, using For {}
if true then
	for i=1 to 10 {
		if i>5 then 100
		print i
	}
	print "no this"
100	Print "ok"
end if

\\ jump inside if then/end if, using For / Next
if true then
	for i=1 to 10 
		if i>5 then exit for 200
		print i
	Next   ' i variable is optional here
	print "no this"
200	Print "ok"
end if




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

                                                             