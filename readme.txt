M2000 Interpreter and Environment

Version 9.7 Revision 8 active-X

Fix a bug in Select case when call a sub inside a Case without use a block of code
if true then
      select case "any"
      Case "Sequence"
            T()    ' here works before in a block {T()}
      End select
      Print "done"
End if

Sub T()
Print "ok"
End Sub




From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate


https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

https://www.dropbox.com/s/xt30bspw6q9pf5f/M2000language.exe?dl=0

http://georgekarras.blogspot.gr/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

http://m2000.forumgreek.com/

                                                             