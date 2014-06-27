using System;
using System.Collections.Generic;
using System.Text;

namespace SplitScreenVisualizer
{
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        public WindowWrapper(IntPtr ip)
        {
            Handle = ip;
        }

        public IntPtr Handle
        {
            get;
            private set;
        }
    }


}
