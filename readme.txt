M2000 Interpreter and Environment

Version 9.8 Revision 18 active-X

New numeric constant Infinity, We can use -Infinty too.
Print Infinity
     1.#INF
Print VAL("1.#INF")=Infinity
Also comparison operators works with infinity.

A correction in iterators made from each() function.
When we program from a item outside the range of items,
now the iterator not entering the while loop

a=(1,2,3)
' change 4 to 1
b=each(a, 4)
while b
	print "never print that"
end while
Print b  ' nothing print 
\\ print from last to start
b=each(a, -1, 1)
while b
	print "cursor", b^, "item", array(b)
end while
Print b  ' print backwards
Print b  ' print backwards


There is a newer Info program - a collection of modules.
When fist time run the interpreter do this in M2000 console:
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

                                                             