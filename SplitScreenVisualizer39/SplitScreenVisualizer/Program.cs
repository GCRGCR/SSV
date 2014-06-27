using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SplitScreenVisualizer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
        
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TriggerForm());
        }

    }
}
// Retrieve the working rectangle from the Screen class using the PrimaryScreen and the WorkingArea properties.
  //          System.Drawing.Rectangle workingRectangle = Screen.PrimaryScreen.WorkingArea;

            // Set the size of the form slightly less than size of working rectangle. 
  //          this.Size = new System.Drawing.Size( workingRectangle.Width *3, workingRectangle.Height );

            // Set the location so the entire form is visible. 
  //          this.Location = new System.Drawing.Point(0, 0);