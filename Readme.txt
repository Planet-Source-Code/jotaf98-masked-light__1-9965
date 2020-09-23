	     +---------------------------------+
	     | DRAWMASKEDLIGHT FUNCTION README |
	     +---------------------------------+

		    +-- INTRODUCTION: --+

 Well, since the first "Light" demo project came out so good,
I had to make another one ;)

 This one has a difference: instead of drawing a light with
the specified radius, it draws a light based on a mask image.
The Radius now only controls the size of the glow effect. It's
also just a bit faster (remember, I'm just about to learn
DirectX to enhance these effects even more!).

 This one would be really cool in a graphics design program
to make a special effect with an image or text (see
Mask_2.bmp to see how text looks like with this), but I'm
sure an application for it could be found in a game :)

 Well, it might be too complex for most people, but you don't
need to understand nothing at all to use it. Just call the
function.

 The problem found in the first demo is still here: if you
use it twice in the same image, same position, some weird
colors appear... (To see this, comment RedrawPicture in the
Picture1_Click event.)
 If you know how to solve it, please e-mail me.

		      +-- "BONUS": --+

 1. The glow is drawn from the inside to the outside, like
if it was growing. So maybe there's no need for using a
backbuffer ;)

3. A code snippet is included -- when you exit the program
and it asks you "Visit http://www.planet-source-code.com
now?", click "Yes". Your browser will take you to that page,
no matter what it is! All you need is a .url file (Internet
Link), and the lines you'll find in the "Form_Unload" event!

		    +-- INSTRUCTIONS: --+

 Simply add the .bas file to your project to use it.
 All you need to do in order to activate the effect is to
call this function.

			+-- USAGE: --+

 DrawMaskedLight(Target As Long, Mask As Long, MaskWidth As
Long, MaskHeight As Long, X As Integer, Y As Integer, RedB
As Byte, GreenB As Byte, BlueB As Byte, Radius As Integer)

		      +-- ARGUMENTS: --+

+- Target:
 This is the hDC where the image will be drawn.
(You'll find this as a property of forms and picture
boxes.
 Just know that it refers to the image used in these
controls.)

+- Mask:
 Same as above, but this is the image used as a mask. Only
white areas will be used.

+- MaskWidth:
 The width of the mask image, in pixels.

+- MaskHeight:
 The height of the mask image, in pixels.

+- X:
 The X coordinates of the place where the light will be
drawn (referring to the center).

+- Y:
 The Y coordinates of the place where the light will be
drawn (referring to the center).

+- RedB
 This is the brightness applied to the Red value. Be
aware that high numbers will make the light totally
white!

+- GreenB
 The brightness applied to the Green value.

+- BlueB
 The brightness applied to the Blue value.

+- Radius:
 The radius of the light (distance from the borders of
the mask to the borders of the light), in pixels.

		       +--EXAMPLE: --+

 Check the sample project included for a working demo.
 But the .exe runs much faster than Visual Basic's Run
Mode!

		+--------------------------+
		|  Made by Jotaf98         |
		|  (João F. S. Henriques)  |
		|                          |
		|  E-mail me at            |
		|  jotaf98@hotmail.com     |
		+--------------------------+