using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Pixelator
{
    public partial class PixelatorRibbon
    {
        private void PixelatorRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private async void btnPixelate_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            try
            {
                var openFile = new OpenFileDialog();
                openFile.Filter = "Image Files (*.bmp;*.jpg;*.jpeg;*.png)| *.bmp;*.jpg;*.jpeg;*.png";

                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    var wb = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;

                    await Pixelate.PixelateFile(wb, openFile.FileName);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
            

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}
