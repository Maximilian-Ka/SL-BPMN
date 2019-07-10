using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace Security_Visio_AddIn
{
   
    public partial class Automation
    {
        
        private void Automation_Load(object sender, RibbonUIEventArgs e)
        {
            LoadImages();
        }
        private void LoadImages()
        {
            button1.ShowImage = true;
            button2.ShowImage = true;
            button1.Image = Properties.Resources.Image2;
            button2.Image = Properties.Resources.Image1;
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OnButton1Clicked();
        }


        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OnButton2Clicked();
        }
    }
}
