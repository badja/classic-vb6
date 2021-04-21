                         ----------------------
                         GTA Wave - Version 3.0
                            31 October 1999
                         ----------------------

                            by Adrian Grucza

                   http://gta.telefragged.com/gtawave
                     http://gta.stomped.com/gtawave
                          gtawave@hotmail.com

                          GTA Wave is freeware
                          --------------------

WHAT IS GTA WAVE?
========================================================================
GTA Wave is a program that lets you listen to, modify, and save sounds
from the game "Grand Theft Auto" and its sequel, "GTA2".


INSTALLING GTA WAVE VERSION 3.0
========================================================================
If you have an earlier version of GTA Wave, it is recommended that you
uninstall it before installing version 3.0. To uninstall it, open the
Control Panel and choose Add/Remove Programs.

Please make sure you complete the Backup Wizard(s) that appear when you
first run GTA Wave, even if you have already made backups. Read the next
section for more information.

New features in version 3.0 are listed at the end of this file.


RUNNING GTA WAVE
========================================================================
To run the program, click on GTA Wave, under Programs in the Start menu.

The first time you run GTA Wave 3.0, the GTA Backup Wizard and/or the
GTA2 Backup Wizard will appear. You should complete these Wizards even
if you have already made backups of the GTA/GTA2 sound files. It enables
you to restore the original sounds individually or all at once.


HOW TO USE GTA WAVE
========================================================================

Opening a GTA sound file
------------------------
Click the Open button (or click Open in the File menu).

To open a GTA sound file, navigate to the folder where GTA is installed,
look in the GTADATA\AUDIO folder, and open one of the SDT files.

To open a GTA2 sound file, navigate to the folder where GTA2 is
installed, look in the DATA\AUDIO folder, and open one of the SDT files.

Playing a sound
---------------
If Auto Play is switched on, then sounds are played when clicked. Auto
Play can be turned on by pressing the toolbar button with a picture of a
cassette.

Sounds can also be played by clicking Play or Play Loop from either the
Sound menu, the right-click popup menu, or the toolbar. Clicking Play
will play the sound once; clicking Play Loop will play the looping part
of the sound repeatedly (see Changing the loop start/end points, below).

Editing a sound
---------------
Editing a sound opens it in either your default Wave File editor
(usually Sound Recorder) or an editor of your choice. You can specify
which editor to use in the Options dialog box (see Options, below). When
you have finished editing the sound, close the editor and you will be
asked if you want to keep the changes.

Clearing sounds
---------------
Clearing a sound reduces its size to zero. It does not remove the sound
from the list, so you can still modify the sound later. You can clear
multiple sounds at once by selecting more than one.

Changing pitch
--------------
Changing the pitch of a sound just changes its sample rate. The actual
sound data does not change. You can change the pitch of multiple sounds
by selecting more than one.

For some sounds, changing the sample rate has no effect on that sound in
the game. For example, car engines and the GTA radio vocals will not
sound any different. This is because for these sounds, the game
disregards the sample rate in the file and uses its own sample rate. In
this case you must use a sound editor to change the pitch. In Sound
Recorder this is done by choosing Increase Speed or Decrease Speed.

Changing pitch variation (GTA2 only)
------------------------------------
Many sounds, such as car collisions, bullets hitting objects, and
footsteps are all played at slightly different pitches during the game.
In GTA, the extent to which the pitches vary is hard-coded into the
program itself. In GTA2, this information is included in the sound
files, so you can control the pitch variation range of each sound.

The pitch variation range is specified by a ± (plus or minus) value, in
Hz. The playback pitch could be anywhere within the range defined by the
sample rate of the sound and this ± value. For example, a
22,050 Hz ± 1,000 Hz sound could be played back at a sample rate
anywhere between 21,050 Hz and 23,050 Hz.

If the Random Pitch Variation option is turned on (look in the Play menu
or on the toolbar), then GTA Wave will playback sounds at a random pitch
within the pitch variation range.

Changing loop start/end points (GTA2 only)
----------------------------------------
In GTA sound files, looping sounds are always played from start to
finish repeatedly. With GTA2 sounds, you can specify the start and end
points for looping. For example, the electrocution sound starts out with
a loud electrical buzz, which turns into a softer sizzling sound. Only
the latter sizzling part of the sound is looped, so as not to repeat the
loud buzz every time the sound loops.

When GTA/GTA2 plays a looping sound, it plays from the start of the
sound until it reaches the loop end point, then plays the sound data
between the loop start and end points repeatedly.

The loop start and end points are specified by numbers (measured in
bytes from the start of the sound) which represent the start and end
points of the looping section. For example, the electrocution sound is
56,576 bytes long, has a loop start point of 30,206 bytes, and a loop
end point at the end of the sound. When a person is electrocuted, the
first 30,206 bytes will play first, followed by the remaining 26,370
bytes looping repeatedly.

The Play and Play Loop commands allow you to play the sound through
once, or play just the looping part of the sound repeatedly.

Importing a sound
-----------------
You can only import a sound if its sampling size (8-bit/16-bit) and
number of channels (mono/stereo) match those of the sound you are
importing into.

In GTA, all the sounds are in 8-bit mono format, except for those in
LEVEL000. The sounds in this file are 16-bit, the first three being
stereo.

In GTA2, all the sounds are in 16-bit mono format, except the cop radio
sounds, which are 8-bit. The sounds in FSTYLE come in pairs, one ending
in 'L' and the other in 'R'. In the game, the two sounds are played
simultaneously through the left and the right speakers respectively.

A quick way to import a sound is to drag the file onto the sound you
want to replace, from an Explorer window or the desktop.

Exporting a sound
-----------------
You can export sounds as Wave Files. A quick way to export a sound is to
drag it into an Explorer window or onto the desktop with the mouse. You
can export multiple sounds by selecting more than one and dragging them.

Restoring sounds
----------------
Once you have completed the appropriate Backup Wizard, you can restore
the original GTA or GTA2 sounds individually by choosing Restore. You
can restore multiple sounds by selecting more than one. To restore all
the sounds in the current file, choose Restore All.

Options
-------
You can change the GTA Wave options by choosing Options from the Edit
menu. Here you can change which action a double click performs, your
sound backup folders, the GTA/GTA2 program files, and which sound editor
to use. You should not need to change the backup folder entries. If one
of them is empty, run the appropriate Backup Wizard from the File menu.


TIPS
========================================================================

GTA2 sound files
----------------
The four SDT files in GTA2 contain the following sounds:

FSTYLE.SDT - menu sounds
WIL.SDT    - area 1 sounds
STE.SDT    - area 2 sounds
BIL.SDT    - area 3 sounds

Extra GTA2 sounds
-----------------
In GTA2, there are extra vocal sounds in WAV format inside the
DATA/AUDIO/VOCALS folder. These WAV files can be opened and modified in
any sound editor, but not from within GTA Wave. Therefore, GTA Wave
doesn't backup these sounds, so it's up to you to do so if you wish to
restore them later.

Sorting the sound list
----------------------
You can change the order in which GTA Wave displays the sounds by
clicking on the column headers. To sort the list in ascending order,
click on the category name you want to sort by. For descending order,
click again.

Playing sounds
--------------
You can control how GTA Wave plays sounds using the Play menu or the
corresponding toolbar buttons. For example, click on the AutoPlay button
to play a sound whenever it is clicked.

File size limits
----------------
Grand Theft Auto will exit with an error if the size of one of the
LEVEL???.RAW files exceeds 1 MB (1,048,576 bytes). In GTA2, RAW files
will not load if they exceed 6,100,000 bytes. You can keep an eye on
this file size and how much it is under/over the limit by looking at
the status bar.

If you are over the limit, first sort the sound list in order of
descending size, by clicking on the Size column header twice. There are
two things you can then do:

* Look for large sounds which you don't mind clearing. Some sounds are
  rarely or never heard in the game.

* Reduce the sample rate of large sounds. This is done with your sound
  editor (not with the Pitch command), and reduces the quality of the
  sound. In Sound Recorder, this is done by choosing Properties from the
  File menu, and clicking Convert Now. Make sure you only change the
  sample rate, and not the sampling size (8-bit/16-bit) or the number of
  channels (mono/stereo).

Using Sound Recorder
--------------------
If your default Wave File editor is Sound Recorder, you may notice that
it adds noise to the end of some sounds when you edit them. This is not
a bug in GTA Wave. You will have to trim off the noise manually, import
the sound directly into GTA Wave, or specify an alternative editor in
the Options dialog box (see Options, above).


UNINSTALLING GTA WAVE
========================================================================
Open the Control Panel, go to Add/Remove Programs, select GTA Wave, and
click Add/Remove.


CHANGES IN VERSION 3.0
========================================================================
* GTA2 support
* Change pitch variation (GTA2 only)
* Change loop start/end points (GTA2 only)
* Random pitch variation (GTA2 only)
* External editor support
* Added "Scale pitch variation range(s)" option to the Change Pitch
  dialog box (GTA2 only)
* Removed Loop from the Play menu and added Play Loop to the Sound menu
* Added another GTA Wave URL in the About box
* Changed the Backup Wizard picture
* Changed the GTA Wave icon


CHANGES IN VERSION 2.01
========================================================================
* Fixed a bug which crashed the program when trying to open the Options
  dialog box once a double-click action other than Open Sound was chosen
* Created a Help menu with an item to view the ReadMe.txt file
* Moved the About box from the File menu to the Help menu
* Changed the GTA Wave URL in the About box
* Changed the background colour of the Backup Wizard picture to blue
  instead of the current desktop colour


CHANGES IN VERSION 2.0
========================================================================
* New Explorer style interface:
    sound descriptions from DMA Design and sound information shown
    sound list sorting
    drag and drop importing/exporting
    right-click pop-up menu
    multiple selections
    show/hide toolbar
    sizable window
    Select All/Invert Selection
* Backup Wizard
* Restore original sounds
* Change pitch (sample rate)
* Run GTA from within GTA Wave
* Options:
    double click behaviour
    file locations
* Close current file
* Stop playing current sound
* Preferences saved to disk:
    last open/import/export folder
    options
    show/hide toolbar
    sound playing settings
    window position and size

------------------------------------------------------------------------
Please send any comments, suggestions or bugs to: gtawave@hotmail.com
