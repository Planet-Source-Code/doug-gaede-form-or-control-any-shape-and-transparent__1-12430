This allows you to have a Form or PictureBox shaped exactly like the image you assign to its PICTURE property.  A specified background color, in this case white, that exists in the picture will be made transparent (and non-existent since you can click on objects behind it).  Other code allows you to drag the form around since it doesn't have a title bar.

Make sure you make the Form or PictureBox size just big enough to expose the graphic, unless you intend on exposing an edge or two (if the background color does not match the transparent color; if they match the exposed form will disappear too, but it wastes CPU power).

I have included the code in a class module and a standard module.  The code is the same, but it is best to use modern Object Oriented Programming methods, which dictates using classes (and no standard modules).
