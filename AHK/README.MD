#AHK

##sendClipboard
A quick utility I developed after a co-worker informed me that employees entering information in there weekly timesheets couldn't copy and paste (Some issue with internet explorer, the only browser you could use).
This script simulates typing the characters contained in the current clipboard, which solved his problem.

##ClipQueue
An extension of "sendClipboard" that uses a queue to store contents of the clipboard. So basically you can copy several items into the queue, and then paste them back in order. This was a fun script to implement to learn more about arrays and objects in AHK

##Paste73CharsPerLine (Not used)
Entering operations is a pain in SAP because of the very slow text editor builted into it. Fortunately you can use an alternative and very simple editor that speeds the process up. Unfortunately, you are only allowed 73 CHARACTERS per line (I haven't figured out why) which makes copy and pasting long operations also a pain. This was my work around that splits long text into clusters of 73 characters and pastes them all in the alternative SAP editor.

However, since I developed the "Set Routing" VBA macros there is not much use for this script any longer

##MatrixDrawingPull
This was my attempt at automating pulling drawings from Matrix. I tried at first to use COM elements or even Python Selenium but this was not possible. I had to resort to image scrapping the screen for the proper elements. Overall I decided that this script might not be that helpful and since developing it taught me very little I stopped. (Already familar with image scrapping when I built a AHK runescape bot)