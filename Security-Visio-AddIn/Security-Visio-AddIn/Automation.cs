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
        public delegate void MyEventHandler();
        public event MyEventHandler Button1Clicked;
        private void Automation_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Automation aut = new Automation();
            //this.Button1Clicked();        
            //Console.WriteLine("ndjdsj");
            //test();
        }

        //private void test()
        //{
        //    this.Button1Clicked();
        //}

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
