using System;
using System.Drawing;
using System.Threading.Tasks;

namespace Pixelator
{
    public static class ImageUtilities
    {
        public static readonly Size HighResolutionSize = new Size(1200, 628);
        public static readonly Size MediumRosolutionSize = new Size(600, 314);
        public static readonly Size LowResolutionSize = new Size(300, 157);

        // thanks to: https://stackoverflow.com/a/6501997/1148564
        public static async Task<Bitmap> ScaleImage(Image image, Size size)
        {
            var ratioX = (double)size.Width / image.Width;
            var ratioY = (double)size.Height / image.Height;
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
