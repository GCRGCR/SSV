using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Threading;

using System.IO;
using System.Globalization;
using System.Diagnostics;

using System.Reflection;
//using System.Security.Permissions;


namespace SplitScreenVisualizer
{
    public partial class TriggerForm : Form
    {

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            Trace.WriteLine(m.Msg.ToString() + ": " + m.ToString());
             if (m.Msg == 0x11)   // WM_QUERYENDSESSION
             {
                 this.Close();
             }
        }

        public struct Settings
        {
            public string strRegName ;
            public string strPPath;
            public float fltPWidth ;
            public int intDPort ;
            public Color color1;
            public Color color2;
            public string strValidTime1;
            public bool bRemoteSet;
            public string strMessage1;
            public string strMessage2;
            public byte bytTimeOver;
            public float fFS;
            public byte bytRunMacro;
            public string strPPMacroName;
            public string strMParam1;
            public string strMParam2;
            public string strMParam3;
            public byte bytFullScreenPres;
            public byte bytAutostart;
            public byte ShowSQLData;
            public string SQLServer;
            public string SQLUser;
            public string SQLPassword;
        }

        static Settings stSettings = Init();
  
        float originalPWidth;
        public static int intSWidth, intSHeight, intSplitWidth;
        public static bool onActivated;
        private System.Windows.Threading.DispatcherTimer dispatcherTimer;
        frmLineMessage LineMessageWindow ;
        SQLForm sqlForm;
        const int MSGWNDMargin =35;
        const int SQLMessageHeight=62;
        static bool bDebugMode;

        PowerPoint.Application objApp;
        PowerPoint.Presentations objPresSet;
        PowerPoint.SlideShowWindows objSSWs;
        PowerPoint.SlideShowSettings objSSS;
        PowerPoint._Presentation objPres;

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetParent(IntPtr child, IntPtr newParent);

        static public Settings Init()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            Settings initSettings = new Settings();
            initSettings.strRegName = @"HKEY_CURRENT_USER\SOFTWARE\Hyg\ScreenSplitter\";
            initSettings.strPPath = @"C:\Presentation.ppt";
            initSettings.fltPWidth = 0.75F;
            initSettings.intDPort = 20000;
            initSettings.color1 = Color.FromArgb(200,50, 50, 50);
            initSettings.color2 = Color.FromArgb( 200,50, 50, 50);
            initSettings.strMessage1 = "NA";
            initSettings.strMessage2 = "NA";
            initSettings.strValidTime1 = "";
            initSettings.bytTimeOver = 0;
            initSettings.fFS = 32;
            initSettings.bytRunMacro = 1;
            initSettings.strPPMacroName = "StartSlideShow";
            initSettings.strMParam1 = "Torque Diagram.xlsx";
            initSettings.strMParam2 = "LCC1";
            initSettings.strMParam3 = "200";
            initSettings.bytFullScreenPres = 0;
            initSettings.bytAutostart = 0;
            
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.AssemblyResolve += new ResolveEventHandler(MyResolveEventHandler);
            //currentDomain.UnhandledException += new UnhandledExceptionEventHandler(MyUnhandledExceptionEventHandler);
            //currentDomain.AssemblyLoad += new AssemblyLoadEventHandler(MyAssemblyLoadEventHandler);
            //currentDomain.TypeResolve +=new ResolveEventHandler(HandleTypeResolve);
            //currentDomain.ResourceResolve += new ResolveEventHandler(MyResourceResolveEventHandler);
            //object[] myobj ;
            //myobj=null;
            //System.Reflection.Assembly ass = System.Reflection.Assembly.LoadFrom("Microsoft.Office.Interop.PowerPoint");
            //Type mytype = ass.GetType("Microsoft.Office.Interop.PowerPoint.Application");
            //var t = Activator.CreateInstance(ass.FullName,mytype.Name);
           // object app = Activator.CreateInstance("Microsoft.Office.Interop.PowerPoint", "Microsoft.Office.Interop.PowerPoint.Application");
           //( currentDomain.CreateInstance("Microsoft.Office.Interop.PowerPoint", "Application"));
            return initSettings;
        }

        private static Assembly MyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            //This handler is called only when the common language runtime tries to bind to the assembly and fails.

            Assembly MyAssembly = System.Reflection.Assembly.LoadWithPartialName(args.Name.Substring(0, args.Name.IndexOf(",")));
            if (MyAssembly == null)
             {
                 MessageBox.Show("Assembly not loaded: " + args.Name  + "\n" +"The progam now exit");
                 Application.Exit();
             }
	        //Return the loaded assembly.
            TriggerForm.onActivated = true;
            MessageBox.Show("Required Assembly missing: \n" + args.Name + " \nUsing this assembly:\n" + MyAssembly.FullName);
            return MyAssembly;      
        }

        /*static void MyUnhandledExceptionEventHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception e = (Exception)args.ExceptionObject;
            MessageBox.Show("UnhandledExceptionEventHandler /n" + sender.ToString() + "\n" + args.ToString() + "\n" + e.InnerException.Message);
            Console.WriteLine("MyHandler caught : " + e.Message);
            Console.WriteLine("Runtime terminating: {0}", args.IsTerminating);
        }   */

        static void MyAssemblyLoadEventHandler(object sender, AssemblyLoadEventArgs args)
        {
            Console.WriteLine("ASSEMBLY LOADED: " + args.LoadedAssembly.FullName);
        }

        public TriggerForm()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            InitializeComponent();
        }

        private void TriggerForm_Load(object sender, EventArgs e)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);
            SystemEvents.DisplaySettingsChanged += new EventHandler(SystemEvents_DisplaySettingsChanged);
            LoadSettings();
            SetWindowsSize();
            SetTriggerTimer();
            //this.TopMost = true;

            string[] args = Environment.GetCommandLineArgs();

            // The first commandline argument is always the executable path itself.
            if (args.Length > 1)
            {
                if (Array.IndexOf(args, "/sp") != -1)
                {
                    tsmPresentationStart_Click(null,null);
                }
            }
            
        }

        private void SystemEvents_DisplaySettingsChanged(object sender, EventArgs e)
    {
        Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
        if (tsmPresentationStop.Enabled == true) // recalculate screen
        {
            tsmPresentationStop_Click(null, null);
            SetWindowsSize();
            tsmPresentationStart_Click(null, null);
        }
        else
        {
            SetWindowsSize();
        }
    }

        private void SetWindowsSize()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            WindowSize();
            SlideShowSize();
            UpdateLabels();
            ShowLabels();
            if (ShowSQLData)
            {
                if (sqlForm == null)
                {
                    sqlForm = new SQLForm();
                }
                sqlForm.Width = intSplitWidth;
                sqlForm.Height = SQLMessageHeight;
                sqlForm.Location = new System.Drawing.Point(0, intSHeight - SQLMessageHeight);
                sqlForm.productline = stSettings.strMParam2;
                sqlForm.sqlreadtime.Ini();
                sqlForm.Show();          
            }
            else
            {
                if (sqlForm != null)
                {
                    sqlForm.Close();
                    sqlForm = null;
                }
            }
        }

        private void WindowSize()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            intSWidth = WinApi.ScreenX; //Screen.PrimaryScreen.WorkingArea.Width;
            intSHeight = WinApi.ScreenY;//Screen.PrimaryScreen.WorkingArea.Height;
            intSplitWidth = System.Convert.ToInt16(System.Convert.ToDouble(intSWidth) * stSettings.fltPWidth);
            this.Size = new System.Drawing.Size(intSWidth - intSplitWidth, intSHeight);
            this.Location = new System.Drawing.Point(intSplitWidth, 0);
        }

        private void TriggerForm_Activated(object sender, EventArgs e)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            if(TriggerForm.onActivated == false)  SSWindowUp();
            TriggerForm.onActivated = false;
            //ShowLabels();
            //SlideShowSize();
        }
 
        private void LoadSettings()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            if (GetReg(stSettings.strRegName, "Path") != "")
            {
                try
                {
                    stSettings.strPPath = GetReg(stSettings.strRegName, "Path");
                    stSettings.fltPWidth = (float)Convert.ToDouble(GetReg(stSettings.strRegName, "Width"));
                    stSettings.intDPort = Convert.ToInt32(GetReg(stSettings.strRegName, "Port"));
                    stSettings.color1 = ColorTranslator.FromWin32(Convert.ToInt32(GetReg(stSettings.strRegName, "Color1")));
                    stSettings.color2 = ColorTranslator.FromWin32(Convert.ToInt32(GetReg(stSettings.strRegName, "Color2")));
                    stSettings.strValidTime1 = GetReg(stSettings.strRegName, "EndTime1");
                    stSettings.strMessage1 = GetReg(stSettings.strRegName, "Message1");
                    stSettings.strMessage2 = GetReg(stSettings.strRegName, "Message2");
                    stSettings.bytTimeOver=Convert.ToByte(GetReg(stSettings.strRegName, "TimeOver"));
                    stSettings.fFS = (float)Convert.ToDouble(GetReg(stSettings.strRegName, "FontSize"));
                    stSettings.bytRunMacro = Convert.ToByte(GetReg(stSettings.strRegName, "RunMacro"));
                    stSettings.strPPMacroName = GetReg(stSettings.strRegName, "MacroName");
                    stSettings.strMParam1 = GetReg(stSettings.strRegName, "MacroParam1");
                    stSettings.strMParam2 = GetReg(stSettings.strRegName, "MacroParam2");
                    stSettings.strMParam3 = GetReg(stSettings.strRegName, "MacroParam3");
                    stSettings.bytFullScreenPres = Convert.ToByte(GetReg(stSettings.strRegName, "FullScreenPresentation"));
                    stSettings.bytAutostart = Convert.ToByte(GetReg(stSettings.strRegName, "Autostart"));
                    stSettings.ShowSQLData = Convert.ToByte(GetReg(stSettings.strRegName, "ShowSQLData"));
                    stSettings.SQLServer = GetReg(stSettings.strRegName, "SQLServer");
                    stSettings.SQLUser = GetReg(stSettings.strRegName, "SQLUser");
                    stSettings.SQLPassword = GetReg(stSettings.strRegName, "SQLPassword");
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.ToString());
                    MessageBox.Show("Error in load settings." + Environment.NewLine 
                                    +  "Please configure the settings!" + Environment.NewLine 
                                    + e.Message.ToString(), "Error"); 
                }
            }
            else
            {
                object ob = "FirstStart";
                tsmSettings_Click(  ob , null);
            }
            StartReceive();
        }

        public static void SaveSettings()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------

            SetReg(stSettings.strRegName, "Path", stSettings.strPPath);
            SetReg(stSettings.strRegName, "Width", Convert.ToString(stSettings.fltPWidth));
            SetReg(stSettings.strRegName, "Port", Convert.ToString(stSettings.intDPort));
            SetReg(stSettings.strRegName, "Color1", ColorTranslator.ToWin32(stSettings.color1).ToString());
            SetReg(stSettings.strRegName, "Color2", ColorTranslator.ToWin32(stSettings.color2).ToString());
            SetReg(stSettings.strRegName, "EndTime1", stSettings.strValidTime1);
            SetReg(stSettings.strRegName, "Message1", stSettings.strMessage1);
            SetReg(stSettings.strRegName, "Message2", stSettings.strMessage2);
            SetReg(stSettings.strRegName, "TimeOver", stSettings.bytTimeOver.ToString());
            SetReg(stSettings.strRegName, "FontSize", Convert.ToString(stSettings.fFS));
            SetReg(stSettings.strRegName, "RunMacro", stSettings.bytRunMacro.ToString());
            SetReg(stSettings.strRegName, "MacroName", stSettings.strPPMacroName);
            SetReg(stSettings.strRegName, "MacroParam1", stSettings.strMParam1 );
            SetReg(stSettings.strRegName, "MacroParam2", stSettings.strMParam2);
            SetReg(stSettings.strRegName, "MacroParam3", stSettings.strMParam3);
            SetReg(stSettings.strRegName, "FullScreenPresentation", stSettings.bytFullScreenPres.ToString());
            SetReg(stSettings.strRegName, "Autostart", stSettings.bytAutostart.ToString());
            SetReg(stSettings.strRegName, "ShowSQLData", stSettings.ShowSQLData.ToString());
            SetReg(stSettings.strRegName, "SQLServer", stSettings.SQLServer.ToString());
            SetReg(stSettings.strRegName, "SQLUser", stSettings.SQLUser.ToString());
            SetReg(stSettings.strRegName, "SQLPassword", stSettings.SQLPassword.ToString());
        }

        private string GetReg(string keyName, string valueName)
        {
            //Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            return Convert.ToString(Registry.GetValue(keyName, valueName, null));
        }

        private static void SetReg(string keyName, string valueName, string value)
        {
            //Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            Registry.SetValue(keyName, valueName, value, RegistryValueKind.String);
        }

        private void UpdateLabels()
        {
            //Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            lblSTrigger.Text = stSettings.strMessage1;
            lblQTrigger.Text = stSettings.strMessage2;
           
            if (stSettings.bytTimeOver==1)
            {
                lblSTrigger.ForeColor = Color.FromArgb(255, 0, 0);
                lblQTrigger.ForeColor = Color.FromArgb(255, 0, 0);
                lblSTrigger.BackColor = stSettings.color1;
                lblQTrigger.BackColor = stSettings.color2;
            }
            else if (stSettings.bytTimeOver == 2)
            {
                lblSTrigger.ForeColor = stSettings.color1;
                lblQTrigger.ForeColor = stSettings.color2;
                lblSTrigger.BackColor = Color.FromArgb(255, 0, 0);
                lblQTrigger.BackColor = Color.FromArgb(255, 0, 0);
            }
            else
            {
                lblSTrigger.BackColor = stSettings.color1;
                lblQTrigger.BackColor = stSettings.color2;
                lblSTrigger.ForeColor = Color.FromArgb(0, 0, 0);
                lblQTrigger.ForeColor = Color.FromArgb(0, 0, 0);
            }
        }

        void ShowLabels()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------

            lblSTrigger.Width = lblSTrigger.Parent.ClientSize.Width;
            lblQTrigger.Width = lblQTrigger.Parent.ClientSize.Width;
            lblSTS.Width = lblSTS.Parent.ClientSize.Width;
            lblQTS.Width = lblQTrigger.Parent.ClientSize.Width;

            lblSTrigger.Height = lblSTrigger.Parent.ClientSize.Height / 2-lblSTS.Height;
            lblQTrigger.Height = lblQTrigger.Parent.ClientSize.Height / 2 - lblQTS.Height;
            lblQTS.Top = lblSTS.Height + lblSTrigger.Height;
            lblSTrigger.Top = lblSTS.Height;
            lblQTrigger.Top = lblSTS.Height + lblSTrigger.Height + lblQTS.Height;

            lblSTrigger.Font =new Font("Microsoft Sans Serif",stSettings.fFS);
            lblQTrigger.Font = new Font("Microsoft Sans Serif", stSettings.fFS);
            lblQTrigger.TextAlign = ContentAlignment.MiddleCenter;

            lblSTrigger.Visible = true;
            lblQTrigger.Visible = true;

            lblDebug.MaximumSize = new System.Drawing.Size(this.Width, 0);
            lblDebug.Visible = bDebugMode;
        }
        
        private Boolean CheckEndTime()  
        { 
            //Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            try
            {
                if (Convert.ToDateTime(stSettings.strValidTime1)<DateTime.Now)
                {
                    stSettings.color1 = Color.FromArgb(200, 50, 50, 50);
                    stSettings.color2 = Color.FromArgb(200, 50, 50, 50);
                    stSettings.bytTimeOver = 0;
                    stSettings.strMessage1 = "NA";
                    stSettings.strMessage2 = Utilitys.getHostIP();

                    lblQTrigger.Font = new Font("Microsoft Sans Serif", 20);
                    lblQTrigger.TextAlign =  ContentAlignment.BottomCenter;
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                stSettings.strValidTime1 = DateTime.Now.ToString();
                return true;
            }
        }

        private void SetTriggerTimer()
        {
            //  DispatcherTimer setup
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            if (dispatcherTimer == null)
            {
                dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
                dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
                dispatcherTimer.Interval = new TimeSpan(0, 0, 2);
            }
            dispatcherTimer.Start();
        }

        //  System.Windows.Threading.DispatcherTimer.Tick handler 
        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            //Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            if (CheckEndTime())
            {
                //dispatcherTimer.Stop();
                UpdateLabels();
            }
            if (sqlForm != null)
            {
                sqlForm.SQLFormRefresh(stSettings.strMParam2);
                if (bDebugMode) //debug
                {
                    lblDebug.Text = "SQL Read Time:" + sqlForm.sqlreadtime.strActualTime + 
                        "    Max:" + sqlForm.sqlreadtime.strMaxTime + 
                        "    Date:" + sqlForm.sqlreadtime.strMaxDate + (char)13 + 
                        "Query:" + sqlForm.sqlreadtime.strRemark1 + (char)13+ 
                        "Last Error:" + sqlForm.sqlreadtime.strRemark2;
                }
            }
        }

        private void StartReceive()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);
            Thread current = Thread.CurrentThread;
            if (current.Name == null) { current.Name = "Main-Thread"; }
            if (!Network.ReciveEnabled)
            {
                Thread NetThread = new Thread(new ParameterizedThreadStart(Network.ServerStart));
                NetThread.IsBackground = true;
                NetThread.Name = "Network thread";
                NetThread.Start( stSettings.intDPort);
                Network.ReceiveMessageEvent += new EventHandler(MSG_Recived);
            }
        }

        private void MSG_Recived(object sender, EventArgs e)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            if (this.InvokeRequired)
            {
                this.Invoke(new EventHandler(MSG_Recived), new object[] { sender, e });
                return;
            }

            if (Network.Message.bRemoteSet == true)
            {
                Network.Message.bRemoteSet = false;
                if (Network.Message.PWidth != 0) { fltWidth = (float)Network.Message.PWidth; }
                if (Network.Message.iFS != 0) { FSSeting = (float)Network.Message.iFS; }
                if (Network.Message.rxPort != 0) { stSettings.intDPort = Network.Message.rxPort; }
                if (2 < Network.Message.strPPath.Length) {stSettings.strPPath = Network.Message.strPPath;}
                SaveSettings();
                tsmSettings_Click(null,null);
                //return;
            }
            else if (Network.Message.Valid)
            {
                Network.Message.Valid = false;
                stSettings.color1 = Network.Message.bg1Color;
                stSettings.color2 = Network.Message.bg2Color;
                stSettings.strMessage1 = Network.Message.strText1;
                stSettings.strMessage2 = Network.Message.strText2;
                stSettings.strValidTime1 = Convert.ToString(Network.Message.EndDate1);
                stSettings.bytTimeOver = Network.Message.byteTimeOver;
                SaveSettings();
                SetTriggerTimer();
                //UpdateLabels();
                SetWindowsSize();
            }

            else if (Network.Message.bLineMessage == true)
            {
                Network.Message.bLineMessage = false;
                
                if (LineMessageWindow == null)
                {
                    LineMessageWindow = new frmLineMessage();
                }
                if (LineMessageWindow.IsDisposed)
                {
                    LineMessageWindow = null;
                    LineMessageWindow = new frmLineMessage();
                }

                WinApi.SearchedDisplay showDisplayProperty = new WinApi.SearchedDisplay();
                showDisplayProperty.uintSearchedDisplay = Network.Message.showOnPrimary;
                WinApi.DisplayProperty(ref showDisplayProperty);

                LineMessageWindow.Width = showDisplayProperty.right - showDisplayProperty.left - MSGWNDMargin * 2;
                LineMessageWindow.Height = showDisplayProperty.bottom - showDisplayProperty.top - MSGWNDMargin - SQLMessageHeight;
                LineMessageWindow.Location = new System.Drawing.Point(showDisplayProperty.left + MSGWNDMargin, showDisplayProperty.top + MSGWNDMargin);

                LineMessageWindow.strMSG = Network.Message.strLineMessage;
                LineMessageWindow.OnTime = Network.Message.intLineMessageTime;
                LineMessageWindow.intMSGFontSize = Network.Message.intLineMessageFS;
                LineMessageWindow.SetFormProperty();
                if(0<LineMessageWindow.OnTime)LineMessageWindow.Show();
                return;
            }
            this.Activate();
        }

  /*      delegate void SetTextCallback(string text);
        private void SetText(string text)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            // InvokeRequired required compares the thread ID of the 
            // calling thread to the thread ID of the creating thread. 
            // If these threads are different, it returns true. 
            if (this.lblSTrigger.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.lblSTrigger.Text = text;
            }
        }                  */

        private void PresentationStart()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            try
            {
                objApp = new PowerPoint.Application();
                objApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                objPresSet = objApp.Presentations;
                objPres = objPresSet.Open(stSettings.strPPath, Microsoft.Office.Core.MsoTriState.msoFalse, // MsoTriState ReadOnly,
                Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);    //MsoTriState Untitled,MsoTriState WithWindow
            }
            catch (Exception e)
            {
                this.Activate();
                MessageBox.Show(e.Message);
                tsmPresentationStop_Click(null, null);
                return;
            }

            //objSlides = objPres.Slides;

            //Run the Slide show
            objSSS = objPres.SlideShowSettings;
            objSSS.ShowType = Microsoft.Office.Interop.PowerPoint.PpSlideShowType.ppShowTypeSpeaker;
            objSSS.LoopUntilStopped = Microsoft.Office.Core.MsoTriState.msoTrue;
            
            if (stSettings.bytRunMacro == 0 || stSettings.strPPMacroName=="")
            {
                objSSS.Run();
            }
            else
            {
                object[] oRunArgs = new Object[] { "'" + objPres.Name + "'!" + stSettings.strPPMacroName, stSettings.strMParam1, stSettings.strMParam2, stSettings.strMParam3 };
                try
                {
                    objApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, objApp, oRunArgs);
                    originalPWidth = objPres.SlideShowWindow.Width;
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    MessageBox.Show(e.Message + "\n" + "\n" + e.Source + "\n" + e.StackTrace);
                    return;
                }
                catch (Exception e)
                {
                    this.Activate();
                    MessageBox.Show(e.Message + "\n" + "\n" + e.InnerException.Source + "\nn" + e.InnerException.Message);
                    return;
                }
            }

            //handle
            //IntPtr hwnd = new IntPtr(objPres.SlideShowWindow.HWND);
            //WindowWrapper handleWrapper = new WindowWrapper(hwnd);
            //SetParent(handleWrapper.Handle, this.Handle);
            //this.Visible = true;
            if (stSettings.bytFullScreenPres==0)
            {
                SlideShowSize();
                //WinApi.HideTray();
                Taskbar.Hide();
            }
        }

        private void SlideShowSize()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            try
            {
                if (objPres != null && objPres.SlideShowWindow != null)
                {
                    //objPres.SlideShowWindow.Height = this.Height;
                    objPres.SlideShowWindow.Width = originalPWidth * stSettings.fltPWidth;
                    objPres.SlideShowWindow.Top = 0;
                    objPres.SlideShowWindow.Left = 0;
                    //objPres.SlideShowWindow.IsFullScreen = true; 
                }
            }
            catch (Exception e)
            {
                tsmPresentationStop_Click(null,null);
            }
        }
        
        private void PresentationStop()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            objSSS = null;
                try
                {
                    objPres.Close();
                }
                catch (System.NullReferenceException)
                {
                }
                catch(Exception objException)
                {
                    //MessageBox.Show("Error"+ objException.StackTrace, "Error"); 
                }
            objPres = null;
            objPresSet = null;
            //WinApi.ShowTray();
            Taskbar.Show();
        }

        private void tsmPresentationStart_Click(object sender, EventArgs e)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            tsmPresentationStart.Enabled = false;
            tsmPresentationStop.Enabled = true;
            PresentationStart();
            this.Activate();  //not working under Win7
            SSWindowUp();     //and need this line
        }

        private void tsmPresentationStop_Click(object sender, EventArgs e)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            
            tsmPresentationStart.Enabled = true;
            tsmPresentationStop.Enabled = false;
            PresentationStop();
            this.Activate();
        }

        private void tsmSettings_Click(object sender, EventArgs e)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            Network.ReceiveMessageEvent -= new EventHandler(MSG_Recived);
            Network.ServerStop();
            if (sender != null)
            {
                frmSettings fSettings = new frmSettings();
                if (sender.ToString() == "FirstStart")
                {
                    fSettings.StartPosition = FormStartPosition.CenterScreen;
                }
                fSettings.ShowDialog();
            }
            StartReceive();
            SetWindowsSize();
        }

        private void tsmExit__Click(object sender, EventArgs e)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            this.Close();
        }

        private void SSWindowUp()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            //this.Activate();
            try
            {
                if (objPres != null && objPres.SlideShowWindow.Active == Microsoft.Office.Core.MsoTriState.msoFalse)
                {
                    objPres.SlideShowWindow.Activate();
                    //bool a = objPres.SlideShowWindow.Active;
                }
            }
            catch 
            {
                if (tsmPresentationStop.Enabled)
                {
                    tsmPresentationStop_Click(null, null);
                }
            }
        }

        private void TriggerForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            PresentationStop();
            if  ( objApp != null )
            {   objApp.Quit();
                objApp = null;
            }
            Network.ServerStop();
            Thread.Sleep(100);
        }

        public static float fltWidth
        {
            get 
            {
                return stSettings.fltPWidth * 100;
            }
            set
            {
                if (value < 5) value = 5;
                if (value > 90) value = 90;
                stSettings.fltPWidth = value / 100; 
            }
        }

        public static string pathSetting
        {
            get { return stSettings.strPPath; }
            set { stSettings.strPPath = value; }
        }

        public static string macroName
        {
            get { return stSettings.strPPMacroName; }
            set { stSettings.strPPMacroName = value; }
        }

        public static bool runMacro
        {
            get 
            {
                if (stSettings.bytRunMacro==1)
                {
                    return true;
                }
                else
                { 
                    return false; 
                }
            }
            
            set 
            { 
                if (value == true )
                {
                    stSettings.bytRunMacro = 1;
                }
                else
                {
                    stSettings.bytRunMacro = 0;
                }
            }   
        }

        public static string macroParam1
        {
            get { return stSettings.strMParam1 ; }
            set { stSettings.strMParam1 = value; }
        }

        public static string macroParam2
        {
            get { return stSettings.strMParam2; }
            set { stSettings.strMParam2 = value; }
        }

        public static string macroParam3
        {
            get { return stSettings.strMParam3; }
            set { stSettings.strMParam3 = value; }
        }

        public static bool FullScreenPres
        {
            get
            {
                if (stSettings.bytFullScreenPres == 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            set
            {
                if (value == true)
                {
                    stSettings.bytFullScreenPres = 1;
                }
                else
                {
                    stSettings.bytFullScreenPres = 0;
                }
            }
        }

        public static bool Autostart
        {
            get
            {
                if (stSettings.bytAutostart == 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            set
            {
                if (value == true)
                {
                    stSettings.bytAutostart = 1;
                }
                else
                {
                    stSettings.bytAutostart = 0;
                }
            }
        }

        public static string PortSeting
        {
            get
            {
                return Convert.ToString(stSettings.intDPort);
            }
            set
            {
                stSettings.intDPort = Convert.ToInt32(value);
            }
        }

        public static float FSSeting
        {
            get
            {
                return stSettings.fFS;
            }
            set
            {
                if (value < 15) value = 15;
                if (value > 45) value = 45;
                stSettings.fFS = (value);
            }
        }

        public static bool ShowSQLData
        {
            get
            {
                if (stSettings.ShowSQLData == 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            set
            {
                if (value == true)
                {
                    stSettings.ShowSQLData = 1;
                }
                else
                {
                    stSettings.ShowSQLData = 0;
                }
            }
        }

        public static string SQLServer
        {
            get { return stSettings.SQLServer; }
            set { stSettings.SQLServer = value; }
        }

        public static string SQLUser
        {
            get { return stSettings.SQLUser; }
            set { stSettings.SQLUser = value; }
        }

        public static string SQLPassword
        {
            get { return stSettings.SQLPassword; }
            set { stSettings.SQLPassword = value; }
        }

        public static bool debugMode
        {
            get { return bDebugMode; }
            set { bDebugMode = value; }
        }

        private void tsmAbout_Click(object sender, EventArgs e)
        {
            AboutBox frmA = new AboutBox();
            frmA.ShowAboutBox();
        }

    }

	public class AboutBox : Form 
    {
        public String sTitle;
        public String sProductName;
        public String sDescription;
        public String sCompanyName;
		public String sAppName;
		public String sVersion;
        public String sCopyrightHolder;
		public Boolean HasBitmap;

		public AboutBox()
        {
            sTitle = Title;
            sProductName = ProductName;
            sDescription = Description;
            sCompanyName = CompanyName;
            sAppName = Assembly.GetExecutingAssembly().GetName().ToString();
            sVersion = Version.ToString();
            sCopyrightHolder = CopyrightHolder;

            this.ShowInTaskbar = false;
            this.Font = new Font("Open Sans", 9);
        }
        
        public void ShowAboutBox ()
		{
			InitDialog ();
		}

		private void InitDialog ()
		{
			this.ClientSize = new Size (250, 140);
			this.Text = "About";
			this.FormBorderStyle = FormBorderStyle.FixedDialog;
			this.ControlBox		= false;
			this.MinimizeBox	= false;
			this.MaximizeBox	= false;

			Button wndClose = new Button ();
			wndClose.Text = "OK";
			wndClose.Location = new Point (90, 105);
			wndClose.Size = new Size (72, 26);
			wndClose.Click += new EventHandler (About_OK);

			Label wndAuthorLabel = new Label ();
			wndAuthorLabel.Text = "Author:";
			wndAuthorLabel.Location = new Point (5, 5);
			wndAuthorLabel.Size = new Size (72, 24);

			Label wndAuthor = new Label ();
			wndAuthor.Text = sCompanyName;
			wndAuthor.Location = new Point (80, 5);
			wndAuthor.Size = new Size (150, 24);

			Label wndProdNameLabel = new Label ();
			wndProdNameLabel.Text = "Product:";
			wndProdNameLabel.Location = new Point (5, 30);
			wndProdNameLabel.Size = new Size (72, 24);

			Label wndProdName = new Label ();
			wndProdName.Text = sTitle;
			wndProdName.Location = new Point (80, 30);
			wndProdName.Size = new Size (150, 24);

			Label wndVersionLabel = new Label ();
			wndVersionLabel.Text = "Version:";
			wndVersionLabel.Location = new Point (5, 55);
			wndVersionLabel.Size = new Size (72, 24);

			Label wndVersion = new Label ();
			wndVersion.Text = sVersion;
			wndVersion.Location = new Point (80, 55);
			wndVersion.Size = new Size (72, 24);

            Label wndCopyrightLabel = new Label();
            wndCopyrightLabel.Text = "Copyrigh:";
            wndCopyrightLabel.Location = new Point(5, 80);
            wndCopyrightLabel.Size = new Size(72, 24);

            Label wndCopyright = new Label();
            wndCopyright.Text = CopyrightHolder.Replace("&", "&&");
            wndCopyright.Location = new Point(5, 80);
            wndCopyright.Size = new Size(240, 24);
                
            this.Controls.AddRange(new Control[] {
												wndClose,
												wndAuthorLabel,
												wndProdNameLabel,
												wndVersionLabel,
												wndAuthor,
												wndProdName,
												wndVersion,
                                                //wndCopyrightLabel,
                                                wndCopyright
												});
			this.StartPosition = FormStartPosition.CenterParent;
			this.ShowDialog ();
		}

		private void About_OK (Object source, EventArgs e)
		{
            this.Close();
		}

        private static Version Version { get { return Assembly.GetCallingAssembly().GetName().Version; } }

        private static string Title
        {
            get
            {
                object[] attributes = Assembly.GetCallingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title.Length > 0) return titleAttribute.Title;
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        new private static string ProductName
        {
            get
            {
                object[] attributes = Assembly.GetCallingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                return attributes.Length == 0 ? "" : ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        private static string Description
        {
            get
            {
                object[] attributes = Assembly.GetCallingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                return attributes.Length == 0 ? "" : ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        private static string CopyrightHolder
        {
            get
            {
                object[] attributes = Assembly.GetCallingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                return attributes.Length == 0 ? "" : ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        new private static string  CompanyName
        {
            get
            {
                object[] attributes = Assembly.GetCallingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                return attributes.Length == 0 ? "" : ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
	};
};
