M2000 Interpreter and Environment

Version 9.7 Revision 6 active-X

fix a bug when we have same name a global array and a local variable
Example (now k(i)-100 executed normal):
k=2
global k(k)
for i=0 to k-1 {
   k(i)=100
}
Print k()



From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate

https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

http://georgekarras.blogspot.gr/

http://m2000.forumgreek.com/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             