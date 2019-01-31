M2000 Interpreter and Environment

Version 9.7 Revision 9 active-X

Fix a bug in trim$(). Added functions ltrim$() and rtrim$()
trim for chars 32 and 160 (nbsp) for the unicode version
and for char 32 the ansi version  (as byte)
also there are variants for ansi strings:
Locale 1033 ' we can set this for proper ansi conversion
k$=Str$("      12345    ")
\\ a length of 7.5 means 15 bytes
print len(k$)*2=15 ' words so x2 =  bytes
m$=rtrim$(k$ as byte)
print len(m$)*2=11, "*"+chr$(m$)+"*" ' convert to utf16
z$=ltrim$(k$ as byte)
print len(z$)*2=9, "*"+chr$(z$) +"*"
x$=trim$(k$ as byte)
print len(x$)*2=5, "*"+chr$(x$)+"*"

\\ Test for string only from spaces
j$=str$("       ")

print len(rtrim$(j$ as byte))=0
print len(ltrim$(j$ as byte))=0
print len(trim$(j$ as byte))=0

print len(rtrim$("    "))=0
print len(ltrim$("    "))=0
print len(trim$("    "))=0





From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate


https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

https://www.dropbox.com/s/xt30bspw6q9pf5f/M2000language.exe?dl=0

http://georgekarras.blogspot.gr/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

http://m2000.forumgreek.com/

                                                             