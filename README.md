# Pixelator
A C# Excel Add-in to embed picture pixels into the worksheet cells

## How it works

In the Ribbon, under the "Add-Ins" tab, there is a new button that says "Pixelate". When you click it, you will be asked to provide an image file. When you provide the file, the program will resize the photo down some if it's too large and then go through pixel-by-pixel and place the color into each respective cell.

### Why do you resize the photo?

In Excel, there are interesting limitations when it comes to styles. For each workbook, there are 

## Original Inspiration

When I first heard about the 73-year-old Japanese man using Excel to make art ([link here](http://www.demilked.com/73-year-old-excel-paintings-tatsuo-horiuchi/)), I first thought that he was painstakingly coloring each individual pixel a specific color. I had always thought "I could code that" by doing the reverse, which is to take a photo and pixelate it into the Excel cells.

That's what this Add-in does.
