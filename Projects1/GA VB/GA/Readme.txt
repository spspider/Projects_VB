 APE (Advanced Platform Engine)
================================

This document will contain basically everything I decide to write down whilst programming the APE.

First things first, I am going to write all the utility apps as I'm programming the game. The game will run in ModeX, and thus be very fast. What we're looking for here is speed, speed and more speed. I want speed. I love speed. I need - speed! :)

I am going to use my CDXVB class library, but I am going to completely modify the code for use with the APE.

There are a few things the APE NEEDS.
1) A map engine - A way of reading in maps of any size, including bad-guys.
2) In-Game music - Probably MIDI support
3) A sprite engine - CDXVBSprite needs some serious updating
4) A sound engine - Dynamic mixing of sounds (I think CDXVBSound is fine as it is)

And a few things I'll add to the APE anyways.
1) Introduction
2) Some cool effects

Score system will be an integer indicating the score, and an integer indicating how many points need to be added to the score. Every frame, as the points are increased, the number of points to add will be decreased.

I'll start with the map engine.

I have decided that each map will require 2 files. The first file is the CDXVBMap loadable map file. The other file will be read in by the game engine, and it will define where every sprite and baddie and everything goes. All the baddies will require x extents so that it can move from left to right.


20th September 1999
===================
Nearly brought the engine up to the kinda level I want now. As in, I have a menu system set up, a rendering system set up, an input engine set up, and I have the outtro credits-a-scrollin'! :)

But once the sound engine is done, that's only the trimmings of the engine done. I still need to finish my map editor, implement a sprite engine and a few other things. But once that's all done, it's just down to making the bitmaps, doing the sounds, and polishing off the game!


Map files use the same tiles in every map. I might change the maps to be binary, so that I'm not limited to 10 tiles (0 - 9 inclusive). In binary, the limit would be 32,768 tiles, which doesn't sound very limiting to me! :)

Sodding maps... :)

Well, I now got more of the trimmings done, and the game is beginning to look really nice now! Soon it should be sounding very nice too... :)

Got some MIDI music support in there! :) The great thing is that you can use ANY MIDI file on the computer for MIDI music. Basically, with every map file, there is a map definitions file too, this file specifies the MIDI music file to use. Also got intro and outtro MIDI music in there. If you dont want music in your game when you adopt this engine, just delete the files. It's as easy as!

Decided to release code as-is, of course I will do some meja development on this myself, but by making the game open-source, I hope to pull in some other VB developers who can help me, also the VB Internet community will benefit from this code. I will include absolutely -EVERYTHING- required to develop this. Well, everything short of the OS, VB6, and a computer. :) Even including the WHOLE CDXVB game programming library, so if people help, they don't have to download CDXVB, they can just pick up the project exactly where I am right now.

Unzip the files win32.tlb and directx.tlb to c:\windows\system and unzip everything else into a directory. Winzip should add the appropriate paths itself.

If you wish to help me with this huge project, be it graphics, sounds, game code, or anything, please e-mail me thegilb@hotmail.com with your name and email address and what you feel you can contribute to the project.

WARNING:
There is no comments. Oh yeah, and MAKE THE EXECUTABLE AND RUN IT FROM EXPLORER RATHER FROM THAN THE VB ENVIRONMENT. There is a bug at the end of the program that I can't work out, and it crashes VB. If you run the program from VB IT MAY LOCK YOUR COMPUTER. Trust me, it is a VERY annoying bug...