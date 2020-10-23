M2000 Interpreter and Environment

Version 10 revision 1 active-X
1. Remove a bug from Bank() function (banker rounding) which casue a removing the minus sign
2. Added size command for controlbox for user forms (so now we can resize a form by using keyboard)
3. Prevent help form to get resize outside the screen
4. Volume statement works in Window 10, and now has a second parameter to reduce volumen from left or right channel
   Also Volume read only variable now reflect the original volume for the specific run, which show the windows control panel.
5. Sound statement now can use buffer from memory to play sounds.
6. SoundRec statement for recording (program in help file)
7. SoundRec.level  as a read only variable
8. Movie statement fix for Windows 10
9. A new example in Info file, console, or how we can use console (same as cmd.exe console) from M2000, using Win 32 api from M2000 code.
10. Sprites a demo as used for my contribution to a CIE 2020 12th Conference on Informatics in Education

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

                                                             