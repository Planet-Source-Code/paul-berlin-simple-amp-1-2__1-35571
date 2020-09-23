-NOTE----------------------------------------------------------
This program uses DMC2 by IzzySoft (http://www.IzzyOnline.com)
You might have to move DMC2.ocx, Bass.dll & ID3v23x.dll to your
Windows system directory and register them with regsvr32.exe.
-NOTE----------------------------------------------------------


			SIMPLE AMP 1.2
		     by Paul Berlin 2002
		   berlin_paul@hotmail.com
-CONTENTS------------------------------------------------------
   Please vote at PlanetSourceCode if you like the program!

 1. Introduction
 2. Features
 3. Quick Help & Explanations 
 4. Other Credits

-1. Introduction-----------------------------------------------

 Simple Amp is an MP3 player for Windows written in 
 Visual Basic 6.0 Enterprise Edition.

 This program is Freeware. If you modify or inprove this
 program in any way it would be great if you could send it
 to me. Please mention me in your program credits if you
 use any of the code. If you find, bugs and don't want to
 fix them yourself, you can always mail an report to me.

 If you make a skin, PLEASE send it to me!

 You should have an fairly fast computer for the scopes to have
 an decent frames per second. I have tested it on an 166Mhz
 Pentium MMX, and it ran pretty slow. I suppose 300-400 Mhz would
 be enough. If it runs slow, you can always disable the scopes.

-2. Features---------------------------------------------------

 * Supports MP1, MP2 & MP3
 * FULLY skinnable program, you can even change size and move
   controls to new locations.
 * 3 Skins included.
 * Supports editing of ID3v1 & ID3v2 tags, even of multiple
   files at the same time.
 * Playlist with lots of features. Save/Load Playlist, add dir,
   sorting, and much more.
 * 5 different scopes that let you SEE the music.
 * Hotkeys. Lets you control the program from the keyboard
   whenever you want (you don't have to have the program in
   focus). Read Quick help for more info.
 * Pitch Control. You can change the pitch of the music, realtime!
 * Advanced soundcard controls. Select mixing quality (Hz, 8/16 bits
   Stereo/Mono), select music buffer size & pan the sound.
 * You can minimize the program to the Systray.
 * The program can be always ontop.
 * The windows snaps to screen edges.

 And lots of more minor features that you can find for yourself.

-3. Quick Help & Explanations-----------------------------------

 Most of the features are easy to understand, but some may be a
 bit harder.

 * Hotkeys
   
   Hotkeys are something I haven't seen in any other program.
   They let you control Simple Amp from for example your keyboard
   or icons in your taskbar toolbars.
   
   It works like this:
   -------------------
   In Simple Amp's settings you can set up what each hotkey will
   do. There are 5 hotkeys in total. In 'action when active' you
   set what will happen for this hotkey when Simple Amp IS started.
   You can for example let it Pause/Play the music. In 'action when
   inactive' you set what will happen when Simple Amp ISN'T started.
   You can for example let it start Simple Amp or start any other 
   program.
   
   Now, when you have set up the hotkeys, you can use them with the
   corresponding hotkey.exe file. To use hotkey 1, you start
   hotkey1.exe and so on.
   
   There are many ways to use hotkeys, you can for example, if you
   have an keyboard with programmable buttons, let one of them start
   an hotkey exe-file that in turn, plays/pauses the music.
   Or you can put icons for each hotkey in an toolbar next to your
   systray, or even at the desktop or in the start menu.

   I have an Microsoft Internet Keyboard, which have two programmable
   buttons. I have the first one set to Start Explorer when Simple Amp
   is inactive and Play/Pause when Simple Amp is active. The second
   buttons starts Simple Amp when it is inactive and plays the next song
   when Simple Amp is active.

   I use them all the time, I have music on ALWAYS (except when
   playing certain games). It's really great when you are in an 
   application which runs fullsceen or hides the taskbar.

 * Using Add/Dir in the playlist
   This can be slow at times, when you add many mp3s, because 
   ID3 tags and other data are read from the files. You can however 
   turn this option off in the settings window to speed up.

 * Editing skins
   To learn how to make skins, check the ini-files for each skin 
   included. They are pretty easy to understand. Just create an new 
   ini-file and images to add an new skin. Everything in the main window 
   and playlist window can be skinned. You can even move and change the 
   size of everything. If you do not want an label to be in your skin, 
   for example the label that shows mp3 kbps & khz, just move it outside 
   the main window or make it an 0x0 size.

 * If you are having problems with your soundcard you could run
   Simple Amp with the -dev switch, it makes the program write
   found soundcard info to device.txt in the program folder.

-3. Other Credits-----------------------------------------------

 I have not written all classes, modules and user controls used
 in Simple Amp myself. Here are credits for them

 Using DMC2 by IzzySoft (http://www.IzzyOnline.com)

 PicScroll & PicVScroll controls by ACP Software

 ID3v23x DLL by Glenn Scott, which I modified to speed up when
 reading tags which contains images or other unsupported
 ID3v2 headers.

 The API filesearch classes clsFile.cls & colFiles.cls and
 the modules mAPIConstants.mod & ts.mod was written by
 The Frog Prince.

 The browse for folder modules modBrowseForFolder.mod &
 modCheck.mod was written by Max Raskin.
 
 The Systray class clsSysTray.cls was written by
 Martin Richardson.
 
 Also, Peeter Puusemp jr. wrote something I used, but forgot
 what it was... sorry!

 I lost track of who wrote modID3v1.mod, which I modified.

 Beta testing by by close friend, Rudi Nilsson.

 Thanks!

----------------------------------------------------------------   