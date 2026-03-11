# Pixelator

A C# Excel Add-in that converts any image into pixel art by mapping 
each pixel's color to its corresponding worksheet cell.

![Pixelator](https://www.spreadsheetsmadeeasy.com/static/39c011ef0f9fe0e6d9a485c6eaf1e3cc/4fa42/pixelator_pic.png)

## See it in Action

![Pixelator in Action](https://www.spreadsheetsmadeeasy.com/ae551af4ef6d4baf960d8c13deb381d7/Pixelate-Image.gif)

[Related Blog Post](https://www.spreadsheetsmadeeasy.com/pixelate-images-in-excel-cells/)

## How It Works

1. Open Excel and navigate to the **Add-Ins** tab in the Ribbon
2. Click the **Pixelate** button
3. Select an image file
4. Pixelator processes the image pixel-by-pixel, coloring each cell to match

> **Note:** Large images are automatically resized to prevent Excel's 
> ["Too many different cell formats"](https://support.microsoft.com/en-us/kb/213904) error.

## Getting Started

> ⚠️ You must create a [signing key](https://msdn.microsoft.com/library/ms247123(v=vs.100).aspx) 
> for the assembly before building.
```bash
git clone https://github.com/joseph4tw/pixelator
# Open in Visual Studio
# Create a signing key (Project Properties > Signing)
# Build and run
```

## License

MIT
