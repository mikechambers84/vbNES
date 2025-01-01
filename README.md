# vbNES v0.1.5

This is a Nintendo Entertainment System emulator that I wrote in Visual Basic 6 back in 2014, and occasionally still hack on a little bit. VB6 obviously is a terrible choice for almost anything these days, but I grew up with it and still find it fun to see what I can do with it now and then.

It supports a decent number of mappers. There are bugs. The compatibility rate isn't great, but it does play a lot of the NES library just fine. The sound emulation also isn't perfect by any means, but most games seem to sound fine.

It has built-in game genie support using the official Game Genie ROM, since Galoob put it into the public domain many years ago. There is some semi-working netplay support. I don't remember exactly what state it's in, but I recall it has issues. It kind of worked on LAN, but don't even bother trying over the internet. I may revisit that part of the code some day.

### HOW TO USE

You need to run the **resreg.bat** script from an ADMINISTRATOR command prompt. This will register a few necessary components with Windows. Then you can run **vbNES.exe** and start playing games.

There are four quick-save slots and the keyboard shortcut for these are F1 to F4. Then F5 to F8 quick-loads them again.

**There is an annoying bug that I need to fix where if you are playing a game, then load a new ROM to play another game, the sound emulation can be messed up. If this happens, close the emulator and restart it before playing the new game.**

### DEFAULT KEY MAPPING
Z = B button
X = A button
Enter = Start button
Right shift = Select button
Arrow keys = Directional pad

Have fun!

Written by Mike Chambers
https://github.com/mikechambers84/

### SCREENSHOTS
![Super Mario Bros 3](https://github.com/mikechambers84/vbNES/blob/main/screenshots/01.png?raw=true)![Mega Man 2](https://github.com/mikechambers84/vbNES/blob/main/screenshots/02.png?raw=true)![Contra](https://github.com/mikechambers84/vbNES/blob/main/screenshots/03.png?raw=true)![The Legend of Zelda](https://github.com/mikechambers84/vbNES/blob/main/screenshots/04.png?raw=true)![Ninja Gaiden](https://github.com/mikechambers84/vbNES/blob/main/screenshots/05.png?raw=true)![Teenage Mutant Ninja Turtles 2](https://github.com/mikechambers84/vbNES/blob/main/screenshots/06.png?raw=true)


### VERSION HISTORY

0.1.5 - Many little fixes. Much improved audio emulation. Compatibility improvement.

0.1.3 - Pretty much fixed MMC3 for most games. Added netplay support. Only seems reliable on LAN for now. Kind of.

0.1.1 - Fixed bug where cancelling a load state dialog would crash the emulator.
      - Added support for 1x, 2x, 3x scaling selection.

0.1.0 - Initial release.

