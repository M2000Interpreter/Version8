M2000 Interpreter and Environment

Version 9.9 active-X

1. Remove of a hard finding bug, which in some OS leave tmp files in %temp% folder. The bug was an abnormal initialization of multimedia player. This bug cause the system to hold the tmp file until the player end. This problem not happen in Windows 8, perhaps is a bug of the external lib, for specific OS.

2. Reorganize the M2000 MessageBoxet, so now we don't get any crash at save.

3. Commands LOAD, SAVE and END check a "dirty" flag to ask for saving. Also REMOVE (with no arguments) now ask before remove the last module/function.

4. Speed improved on calculations (expression evaluation).

5. List of COM object's members, return the type of parameters, and if they are IN, OUT, IN OUT, also the return type if the member return value.

6. New objects., SOCKET, DOWNLOAD, CLIENT for use TCP/IP. Info.gsb have the Down3, a module example to show how to download asynchronous three files, using the DOWNLOAD object.

7. (This break compatibility  with previous versions). Com Events in previous versions internal passed by reference. So all variables in event service function have & before. Now only for those parameters which are ByRef we need to use &. Before we use an event we can just show the stack, and leave it as is, using the Stack statement without parameters. So we can see if we have values or references, because references in M2000 are strings -they are weak references-with name, and for events they have EV at the begin).

8. New info.gsb. Best of all is the chess game (without AI) but import and export FEN strings, and we can play chess with a friend. Also we can replay the game, or move to any previous move. Also there is a Snake board game, only an aytomatic version (4 players played by computer).


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

                                                             