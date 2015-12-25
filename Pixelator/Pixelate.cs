using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace Pixelator
{
    public sealed class Pixelate
    {
        public async static Task PixelateFile(Workbook Workbook, string FileName)
        {
            var image = new Bitmap(FileName);

            // if image is too large, we need to resize it
            if (image.Width > 600 || image.Height > 314)
            {
                var temp = ScaleImage(image, 600, 314);
                // dispose the first handle since we won't be using it
                image.Dispose();

                image = temp;
            }

            var ws = (Worksheet)Workbook.Sheets.Add();
            var usedRange = ((Range)ws.Range[ws.Cells[1, 1], ws.Cells[image.Height + 1, image.Width + 1]]);
            usedRange.ColumnWidth = 0.5;
            usedRange.RowHeight = 4.5;

            for (int x = 1; x <= image.Width; x++)
            {
                for (int y = 1; y <= image.Height; y++)
                {
                    var pixel = image.GetPixel(x - 1, y - 1);

                    ((Range)ws.Cells[y, x]).Interior.Color = pixel;
                }
            }

            image.Dispose();
        }

        // thanks to: https://stackoverflow.com/a/6501997/1148564
        private static Bitmap ScaleImage(Image image, int maxWidth, int maxHeight)
        {
            var ratioX = (double)maxWidth / image.Width;
            var ratioY = (double)maxHeight / image.Height;
            var ratio = Math.Min(ratioX, ratioY);

            var newWidth = (int)(image.Width * ratio);
            var newHeight = (int)(image.Height * ratio);

            var newImage = new Bitmap(newWidth, newHeight);

            using (var graphics = Graphics.FromImage(newImage))
                graphics.DrawImage(image, 0, 0, newWidth, newHeight);

            return newImage;
        }
    }
}
