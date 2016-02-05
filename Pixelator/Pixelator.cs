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
    public static class Pixelator
    {
        public async static Task PixelateFile(Application Application, string FileName)
        {
            var image = new Bitmap(FileName);

            // if image is too large, we need to resize it
            if (image.Width > ImageUtilities.LowResolutionSize.Width
                || image.Height > ImageUtilities.LowResolutionSize.Height)
            {
                var temp = await ImageUtilities.ScaleImage(image, ImageUtilities.LowResolutionSize).ConfigureAwait(false);                

                // dispose the first handle since we won't be using it
                image.Dispose();

                image = temp;
            }

            Workbook wb = Application.Workbooks.Add();
            Worksheet ws = wb.Sheets.Add();
            
            Range usedRange = ws.Range[ws.Cells[1, 1], ws.Cells[image.Height, image.Width]];
            usedRange.ColumnWidth = 0.25;
            usedRange.RowHeight = 2.25;

            for (int x = 1; x <= image.Width; x++)
            {
                for (int y = 1; y <= image.Height; y++)
                {
                    var pixel = image.GetPixel(x - 1, y - 1);

                    ws.Cells[y, x].Interior.Color = pixel;
                }
            }

            image.Dispose();
        }

    }
}
