SnapShot

SnapShot is a simple screen capturing program. It will capture a selected area of the desktop.

After clicking on the Get Snap Shot button move the mouse to the start position of the area you wish to capture. Hold down the left mouse button and move to the end position of selection.  As the mouse moves it will invert the selected area.  The mouse may be positioned more accurately with the arrow keys both with the left button up or down.

Let the left button up to make the selection.  The selections will be saved to frmShots and to the clipboard.  The saved shots may be saved to disk as bitmaps from frmShots.
  
The current position and 19x19 pixel grid around the current position are magnified (x11) in the MagGlass control in the left corner of the screen.  It also displays the current position x,y co-ordinates and the color info for the current pixel.  If the mouse move within 2 pixels of the MagGlass it will switch sides on the screen.

The program is fairly well documented and easy to understand.  It makes extensive use of memory Dcs to do most of the drawing. The only control on the MagBox control is a flat picture box, everything else is drawn to it.

MemoryBmp.cls is useful for my purposes and will be developed further as I make more use of memory Dcs.

Any comments and/or suggestions would be appreciated.


Jim Benvenuti
benj@sympatico.ca  