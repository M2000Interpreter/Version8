M2000 Interpreter and Environment

Version 9.6 Revision 2 active-X

Fix a bug for types for objects for passing by reference:
Group alfa {x=10}
Module AlfaInc (&a as Group) { Print a.x : a.x++}
AlfaInc &alfa   ' 10
AlfaInc &alfa   ' 11
This bug not exist in 9.5 version of early revisions.


From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate

https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

http://georgekarras.blogspot.gr/

http://m2000.forumgreek.com/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             