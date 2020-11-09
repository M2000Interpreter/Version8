M2000 Interpreter and Environment

Version 10 revision 6 active-X
These are serious fixes
1. fix the BASEG example, now work fine (problem introduced when change the way M2000 read properies on com objects, on retrurned objects with value auto property)
2. fix the S1 example (problem introduced when // added as extra remark).
3. Added some functionality to ZipTool (the compressor), so now we can zip to buffer, and unzip a file to buffer. Example Jukebox now has a smaller binary part (Base64), as a zip file in base64, so we get the image from that zip to an expanded one, without using disk operations.
4. A lot of refactoring.


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com


The fist time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)


https://www.dropbox.com/s/30g5oduqt7tzfpm/ca.crt?dl=0

https://www.dropbox.com/s/xt30bspw6q9pf5f/M2000language.exe?dl=0

http://georgekarras.blogspot.gr/

https://github.com/M2000Interpreter/Version9

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             