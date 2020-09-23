
==================================================================================================================
*IMPORTANT NOTICE! In order to run game engine "pak#.pak" files must be in the same folder with application.     *
*		   To run editor you also need editor.ini be in the same folder with editor application.         *
==================================================================================================================

SkaterBoy v1.0 [June.15.2003]
=================
A platform game using DirectX7.
This projects uses PAK file system to store and access resources such as bmps and waves
PAK System is the last step of my same type of submissions (Seacrh PSC as "BMP Packer" and "Ultimate Packer").With this system, developer easily hides his/her game's resources from end-user.Although its not 100 % safe(Nothing is 100 % safe),its pretty straight forward,and fast.As you see if you don't have a PAK file editor such as PakScape you can't change or extract the graphics and sound effects of this game.

Re-release in [June 19 2003]
=================
FIXED:All projects are in same folder,project stable and less complex now now.
ADDED:Textbox for fps limiter changes.(Some users had choppy graphics,and poor performence,I limit fps to 100 in my Celeron 1100 MHz,slower computers may increase textbox's value up to 1000 at which does not limits fps.)
REMOVED:replaced a background which is Ari Feldmans,Now all graphics credit is mine.

How PAK System works
=================
At start up LoadPaks procedure searches application directory for pak#.pak files (where # denotes sequenital numbers from 0 to ....)Loads all pak files contents into a dinamic array
When application needs a file asks PAK system if it's in PAK files,If answer is yes it loads it from there
If no,It looks under the same folder structure in application directory and loads it from there,If not there too then creates a "log.txt" entry ("Error! file xxxxx not found") and quits application.
If a file both exists in pak file and directory,directory will be chosen.

Files:
=================
There are two pak files in this package
pak0.pak:gfx folder,background bitmaps and objects ,sprite bmp files.
pak1.pak:sfx and maps folders,sound effects and map files.

*To extract these files you can use pakextr.exe application in extract_pak folder
rename pakextr.ex to pakextr.exe and run it,But I recommend use of PakScape


About Pak Files:
=========
Pak files are used in some games such as Quake 1-2 ,Half Life etc.
A Pak file is simply an uncompressed zip file,which consists of game resources
such as bitmaps,waves or any other specific game files,in a folder like structure.Pak.cls class handles all pak
loading stuff,to create or edit a pak file I recommend Peter Engström's PakScape.

Features:
=========
1)fast and smooth graphs using DX7
2)almost %70 completed game engine
3)excellent use of pak files ,run-time resource loading
4)If you want to add custom graphs,just add a "gfx" folder or modify pak#.pak files or create a new pak#.pak file

Known issues & Bugs:
=========
1)editor can't load maps in pak files
2)sometimes late refreshing of screen at start up
3)slowing at the edges of background bmps
4)Some bugs at jumping routine

Credits:
=========
Programming,Graphics:Levent Baykan Bayar
Gravity physics help by NDB

Contact:
=========
leventbbayar@operamail.com
Feel free to ask,suggest,comment or threat.

You can use this code in your projects for free as long as you give me a credit.

happy coding.



