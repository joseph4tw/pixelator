# Pixelator

A C# Excel Add-in to embed picture pixels into the worksheet cells

![Pixelator](https://www.spreadsheetsmadeeasy.com/static/39c011ef0f9fe0e6d9a485c6eaf1e3cc/4fa42/pixelator_pic.png)

## See it in Action

![Pixelator in Action](https://www.spreadsheetsmadeeasy.com/ae551af4ef6d4baf960d8c13deb381d7/Pixelate-Image.gif)

[Related Blog Post](https://www.spreadsheetsmadeeasy.com/pixelate-images-in-excel-cells/)

### Why do you resize the photo?

If the image is too large, then it will be resized to fix the ["Too many different cell formats"](https://support.microsoft.com/en-us/kb/213904) error.

## Original Inspiration

When I first heard about the 73-year-old Japanese man using Excel to make art ([link here](http://www.demilked.com/73-year-old-excel-paintings-tatsuo-horiuchi/)), I first thought that he was painstakingly coloring each individual pixel a specific color. I had always thought "I could code that" by doing the reverse, which is to take a photo and pixelate it into the Excel cells.

That's what this Add-in does.

Later - when I actually read the article and looked at the images - I realized that the Japanese man was actually using a combination of things, but mostly Excel Shapes to get the art done. This is much different than I had originally thought. However, I still thought this little app was pretty cool, so I wrote it up for fun.

## If you're going to fork / clone this

Be sure to create a [signing key](https://msdn.microsoft.com/library/ms247123(v=vs.100).aspx) for the assembly before building / running locally.
