using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Diagnostics;
using System.Threading;
using System.Data.SqlClient;
using System.Reflection;
using System.Collections;
  
namespace SplitScreenVisualizer
{
    public partial class SQLForm : Form
    {
        public SQLForm()
        {
            InitializeComponent();
            refreshCounter = refreshSQLCount;
            cTextScroll = new TextSroll(this);
            strSQlCon = "user id=" + TriggerForm.SQLUser +      //linedata;" +
                        ";password=" + TriggerForm.SQLPassword + //robot01;" +
                        ";server=" + TriggerForm.SQLServer +     //HYG-CPS001;" +
                        ";Trusted_Connection=no;" +
                        "database=1982_PG_Csiomor; " +
                        "connection timeout=20";

            Thread SQLDataConnectionThread = new Thread(new ThreadStart(SQLDataConnection));
            SQLDataConnectionThread.IsBackground = true;
            SQLDataConnectionThread.Name = "SQL DataConnection Thread";
            SQLDataConnectionThread.Start();
            ShowStartLabel();
        }

        #region SQL Form Data Tags
        int intTargetBoxNR = 0, intProducedBox = 0;
        string strProdTargPerc;

        const int refreshSQLCount = 5;

        public string productline { get; set; }
        string strSQlCon, strQuer, strLineID;
        int refreshCounter;
        
        SQLData SQLDisplay= new SQLData();
        ManualResetEvent done = new ManualResetEvent(false);

        TextSroll cTextScroll;
        //TextSroll.DisplayText[] arrDisptext ;

        enum fontsize : int{ small=18, big=40 };

        private struct SQLData
        {
            public SQLData(int size)
            {
                strData = new string[size];
                strDescription = new string[size];
                count = size;
                strError = null;
                strTrace = null;
                readretry = 1;
                ErrorCode = -1;
            }
            public string[] strData;
            public string[] strDescription;
            public string strTrace;
            public int count; 
            public string strError;
            public int ErrorCode;

            private int readretry;
            public int ReadRetry
            {   get
                { return readretry; } 
                set 
                {  readretry=value; }
            }
        }

        private string[] Desc = new string[]{
            "PO:", 
            "ProductCode:", 
            "CasesProduced:", 
            "InTransit:", 
            "OnPickPoint:", 
            "Palletized:", 
            "Scrapped:" };

        public struct structTimeMeasure
        {
            public string strActualTime, strActualDate, strMaxTime ,strMaxDate,strRemark1,strRemark2;
            public double dblMaxTime;
            public void Ini()
            {
                strActualTime = "";
                strActualDate = "";
                strMaxTime = "";
                strMaxDate = "";
                strRemark1 = "";
                strRemark2 = "";
                dblMaxTime = 0;
            }
        }
        public structTimeMeasure sqlreadtime;

        #endregion
        //Label uj;
        public void SQLFormRefresh( string productline)
        {
            Debug.Print("Foglalt memória: {0}", GC.GetTotalMemory(false));
            Debug.Print("StToCh:"+cTextScroll.startToChange + " TickToEnd: " + cTextScroll.tickToTickEnd + 
                " TickEndtoP: " + cTextScroll.tickendToPaint + " PainttoEndP: " + cTextScroll.paintToEndPaint + 
                " EndPtoTick: " + cTextScroll.endpaintToTick + " min:" + cTextScroll.mint + " max:" + cTextScroll.maxt + " total:" + cTextScroll.tot);
            cTextScroll.mint = 10000;
            cTextScroll.maxt = 0;

            if (refreshCounter-- <=1)
            {
                refreshCounter = refreshSQLCount;
                done.Set();
            }
        }

        private void SQLDataConnection()
        { 
            TextSroll.DisplayText[] arrDisptext;
            arrDisptext = null;
            string strErrorTrace;
            using (SqlConnection connection = new SqlConnection(strSQlCon))
            {
                while (TriggerForm.ShowSQLData)
                {
                    if (connection.State != ConnectionState.Open)
                    {
                        try { 
                            connection.Open(); }
                        catch (Exception e) {}
                    }
                    Debug.WriteLine("Read SQL");
                    done.Reset();

                    strErrorTrace = "Error1 Line ID Request : ";
                    strQuer = " SELECT LineCode FROM tblLines WHERE (LineDescr ='" + productline + "')";
                    if (ReadSQLRow(strQuer, connection, out SQLDisplay, strErrorTrace))
                    {
                        strLineID = SQLDisplay.strData[0];

                        strErrorTrace = "Error2 Current Assigment Request : ";
                        strQuer =
                            "SELECT QuantityToProduce " +
                            "FROM vw_CurrentAssignments " +
                            "WHERE (LineCode ='" + strLineID + "')";
                        if (ReadSQLRow(strQuer, connection, out SQLDisplay, strErrorTrace))
                        {
                            int.TryParse(SQLDisplay.strData[0], out intTargetBoxNR);

                            strErrorTrace = "Error3 Actual Data Request : ";
                            strQuer =
                                "SELECT ProcessOrder, ProductCode, CasesProduced, CasesInTransit, CasesOnPickPoint, CasesPalletized, CasesScrapped " +
                                "FROM vwWebReport2 " +
                                "WHERE (LineCode = '" + strLineID + "') " +
                                "ORDER BY LineCode";
                            int readAttempt = 0;
                            while (!ReadSQLRow(strQuer, connection, out SQLDisplay, strErrorTrace) && readAttempt < 3)
                            {
                                readAttempt += SQLDisplay.ReadRetry;
                            }

                            if (SQLDisplay.strData != null)
                            {
                                intProducedBox = Convert.ToInt32(SQLDisplay.strData[2]);
                                //strTemp = Convert.ToString((float)intProducedBox / (float)intTargetBoxNR * 100, IFormatProvider );
                                strProdTargPerc = ((float)intProducedBox / (float)intTargetBoxNR).ToString("P");
                                SQLDisplay.strData[2] = SQLDisplay.strData[2] + " / " + intTargetBoxNR + "   " + strProdTargPerc;
                            }
                        }
                    }

                    SQLDataToDisplaytext(ref arrDisptext);

                    cTextScroll.SetText(arrDisptext);
                    if (!cTextScroll.ScrollingRun)
                    {
                        HideStartLabel();
                        cTextScroll.StartScrolling();
                    }

                    Debug.Print("SQL Read End");
                    done.WaitOne();
                }
            }
        }

        

        private void SQLDataToDisplaytext(ref TextSroll.DisplayText[] returnArrDisptext)
        {
            if (SQLDisplay.ErrorCode < 0)   // If NO Error
            {
                returnArrDisptext = new TextSroll.DisplayText[SQLDisplay.strDescription.Count() * 2];

                int c = 0;
                for (int i = 0; i < SQLDisplay.strDescription.Count(); i++)
                {
                    returnArrDisptext[c] = cTextScroll.AddNewDispText(Desc[i]);
                    returnArrDisptext[c].TextWidth = returnArrDisptext[c].TextWidth - 10;
                    //disptext[c].TextString = SQLDisplay.strDescription[i]+":";//.Substring(0,1); */
                    c++;

                    returnArrDisptext[c] = cTextScroll.AddNewDispText(SQLDisplay.strData[i]);
                    returnArrDisptext[c].TextWidth = returnArrDisptext[c].TextWidth + 70;

                    c++;
                }
                cTextScroll.YPos = 0;
                returnArrDisptext[5].TextBackColor = intProducedBox < intTargetBoxNR ? null : Brushes.DarkGoldenrod;
                //ShowText(SQLDisplay.strDescription[intActColumn] + ": " + SQLDisplay.strData[intActColumn], (int)fontsize.big);
            }
            else                            // On Error
            {
                returnArrDisptext = new TextSroll.DisplayText[2];
                returnArrDisptext[0] = cTextScroll.AddNewDispText(SQLDisplay.strTrace);
                returnArrDisptext[1] = cTextScroll.AddNewDispText(SQLDisplay.strError);
                for (int i = 0; i < 2; i++)
                {
                    returnArrDisptext[i].TextFont = new System.Drawing.Font("Microsoft Sans Serif", 18, FontStyle.Bold);
                    cTextScroll.CalculateDisplayTextWidth(returnArrDisptext[i], 50);
                    cTextScroll.CalculateDisplayTextHeight(returnArrDisptext[i]);
                }

                cTextScroll.YPos = this.Height/2 - returnArrDisptext[0].TextHeight/2;
                //ShowText(SQLDisplay.error, (int)fontsize.small, FontStyle.Regular);
            }
        }


        #region Connect to SQL Server and read data
        private bool ReadSQLRow(string queryString, SqlConnection connection, out SQLData dsdReturnData, string trace = "")
        {
            dsdReturnData = new SQLData();
            Stopwatch sw = new Stopwatch();
            TimeMeasure(sw ,ref sqlreadtime);
            try
            {
                //using (SqlConnection connection = new SqlConnection(connectionString))  //{
                SqlCommand command = new SqlCommand(queryString, connection);
                //command.Connection.Open();
                
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    reader.Read();
                    dsdReturnData = new SQLData(reader.FieldCount);
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        dsdReturnData.strData[i] = Convert.ToString(reader[i]);//reader.GetString(i);
                        dsdReturnData.strDescription[i] = reader.GetName(i);
                    }
                    reader.Close();
                    return true;
                }  //}
                
            }
            catch (SqlException err)
            {
                dsdReturnData.strError = err.Message;
                dsdReturnData.strTrace = trace;
                dsdReturnData.ErrorCode = 1;
                return false;
            }
            catch (Exception err)
            {
                dsdReturnData.strError = err.Message;
                dsdReturnData.strTrace = trace;
                dsdReturnData.ErrorCode = 2;
                return false;
            }
            finally
            {
                TimeMeasure(sw,ref sqlreadtime);
                sqlreadtime.strRemark1 = queryString;
                if (dsdReturnData.ErrorCode != 0)
                {
                    sqlreadtime.strRemark2 = dsdReturnData.strError;
                }
            }
        }

        private void ShowStartLabel()
        {
            Label StartLabel = new Label();
            StartLabel.ForeColor = Color.LightGray;
            StartLabel.Font = new Font("Microsoft Sans Serif", 38, FontStyle.Bold);
            StartLabel.Height = this.Height;
            StartLabel.Width = 800;
            StartLabel.Text = "Connect to SQL Server...";
            this.Controls.Add(StartLabel);
            //this.lblSQLData.Text = text;
            //this.lblSQLData.Font = new System.Drawing.Font("Microsoft Sans Serif", (float)fontsize,style);
        }

        private void TimeMeasure(Stopwatch sw, ref structTimeMeasure measure )
        {
            if ( !sw.IsRunning )
            { 
                sw.Start();
            }
            else { 
                sw.Stop();
                measure.strActualDate = DateTime.Now.ToString();//ToShortDateString()+DateTime.Now.ToShortTimeString() ;
                measure.strActualTime = sw.Elapsed.Milliseconds.ToString() + "ms";
                if (sw.Elapsed.TotalSeconds  > 1)
                {
                    measure.strActualTime = Convert.ToInt32(sw.Elapsed.TotalSeconds).ToString() + "s" + measure.strActualTime;
                }

                if (measure.dblMaxTime < sw.Elapsed.TotalMilliseconds)
                {
                    measure.dblMaxTime = sw.Elapsed.TotalMilliseconds;
                    measure.strMaxTime = measure.strActualTime;
                    measure.strMaxDate = measure.strActualDate;
                }
            }
        }

        private delegate void hsl();
        private void HideStartLabel()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new hsl(HideStartLabel));
                return;
            }
            //this.Invoke(new EventHandler(MSG_Recived), new object[] { sender, e });
            this.Controls.Clear();
        }

        private void SQLForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            done.Set();
            //this.SQLDataConnection
        }
        //protected override void Dispose(bool disposing) {        }
    }
    #endregion

    #region  TextSroll class -------------------------------------------------------------
    public class TextSroll
    {
        Stopwatch sw1 = new Stopwatch();
        Stopwatch sw2 = new Stopwatch();
        public long startToChange ,tickToTickEnd, tickendToPaint, paintToEndPaint, endpaintToTick, mint,maxt,tot;

        public TextSroll(Form parentForm)
        {
            dispTimer = new System.Windows.Threading.DispatcherTimer();
            dispTimer.Tick += new EventHandler(dispTimer_Tick);
            parent = parentForm;
            Interval = 25;
            MovePixel = 3;
            InmediateRefresh = true;
            rightToLeft = true;
            Reset();
        } 
        
        
        public class Workdata
        {
            public Workdata()
            {
                TextString = "";
                TextFont = new System.Drawing.Font("Microsoft Sans Serif", 40, FontStyle.Bold);
                TextColor = Color.FromKnownColor(KnownColor.ButtonFace );//FromName(ButtonFace) .FromArgb(255, 255, 255);
                UpdateNeeded = true;
            }

            public string TextString { get; set; }
            public Font TextFont { get; set; }
            public Color TextColor { get; set; }
            public Brush  TextBackColor { get; set; }
            public int TextWidth { get; set; }
            public int TextHeight { get; set; }
            public bool OnDisplay { get; set; }
            public bool UpdateNeeded { get; set; }
        }


        public class  DisplayText : Workdata
        {
            new private bool OnDisplay { get; set; }
            new private bool UpdateNeeded { get; set; }          
        }

        #region Data tags
        Form parent = new Form();
        private System.Windows.Threading.DispatcherTimer dispTimer;

        private DisplayText[] Inputdata;
        private List<Workdata> wrkList = new List<Workdata>();
        private bool TextChanged { get; set; }
        private int firstStringPointer { get; set; }

        private int interval;
        public int Interval
        {
            get { return interval; }
            set
            {
                interval = value;
                dispTimer.Interval = new TimeSpan(0, 0, 0, 0, value);
            }
        }
        public int MovePixel { get; set; }
        public int XPos { get; set; }
        public int YPos { get; set; }
        public bool InmediateRefresh { get; set; }

        private bool leftToRight;
        public bool LeftToRight
        {
            get {return leftToRight; }
            set { 
                leftToRight= value;
                rightToLeft = !value;
            }
        }
        private bool rightToLeft;
        public bool RightToLeft
        {
            get { return rightToLeft; }
            set
            {
                rightToLeft = value;
                leftToRight = !value;
            }
        } 
        
        private bool scrollRunning;
        public bool ScrollingRun
        {
            get { return scrollRunning; }
            set { }
        }
        #endregion

        public TextSroll.DisplayText AddNewDispText(string inText)
        {
            TextSroll.DisplayText retData = new TextSroll.DisplayText();
            retData.TextString = inText;
            CalculateDisplayTextWidth(retData);
            CalculateDisplayTextHeight(retData);

            return retData;
        }

        public void CalculateDisplayTextWidth(TextSroll.DisplayText item, int xCorrection = 0)
        {
            Graphics grfx = parent.CreateGraphics();
            item.TextWidth = (int)grfx.MeasureString(item.TextString, item.TextFont).Width + xCorrection;
        }

        public void CalculateDisplayTextHeight(TextSroll.DisplayText item, int yCorrection = 0)
        {
            Graphics grfx = parent.CreateGraphics();
            item.TextHeight = (int)grfx.MeasureString(item.TextString, item.TextFont).Height + yCorrection;
        }
        
        public void SetText(DisplayText[] input)
        {
            if (input == null){return;}
            Inputdata = input;
            TextChanged = true;
            wrkList.ForEach(v => v.UpdateNeeded = true);
            if(!ScrollingRun ){ChangeText();}
        }

        private void ChangeText()
        {
            bool UpdateCollection = false;

            while (wrkList.Count < Inputdata.Count())
            {
                wrkList.Add(new Workdata());
            }

            while (wrkList.Count > Inputdata.Count())
            {
                UpdateCollection = true;

                int idx = wrkList.FindLastIndex(delegate(Workdata item) { return item.OnDisplay == false; });
                if (idx != -1 && (idx > (Inputdata.Count() - 1)))
                {
                    wrkList.RemoveAt(idx);
                    UpdateCollection = true;
                    if (idx <= firstStringPointer) { firstStringPointer--; }
                }
                else
                {
                    break;
                }
            }

            for (int i = 0; i < Inputdata.Count(); i++)
            {
                if (wrkList.Count >= Inputdata.Count() && (!wrkList[i].OnDisplay || InmediateRefresh))
                {
                    wrkList[i] = (Workdata)Inputdata[i];
                    wrkList[i].UpdateNeeded = false;
                }

                UpdateCollection = UpdateCollection | wrkList[i].UpdateNeeded; //javítani néha out of range exeption
            }
            TextChanged = UpdateCollection;

            // parent.Invalidate ();
            //parent.Refresh();
            //grfx2.Flush(System.Drawing.Drawing2D.FlushIntention.Sync);
        }

        private void dispTimer_Tick(object sender, EventArgs e)
        { //tickToTickEnd, tickendToPaint, paintToEndPaint, endpaintToTick;
            int sleeptime;
            
            int stringwidth = 0;
            int counter = 0;
            int tempPointer = firstStringPointer;
            endpaintToTick = sw1.ElapsedMilliseconds;
            if (TextChanged) ChangeText();
            startToChange = sw1.ElapsedMilliseconds;
            while (stringwidth + XPos<parent.Width && (counter<wrkList.Count))
            {
                if (tempPointer >= wrkList.Count)
                { tempPointer = 0; }
                stringwidth = stringwidth + wrkList[tempPointer].TextWidth;
                wrkList[tempPointer].OnDisplay = true;
                counter++;
                tempPointer++;
            }

            XPos = XPos - MovePixel;

            counter = 0;
            while (counter<wrkList.Count )
            {
                if (XPos + wrkList[firstStringPointer].TextWidth <= 0 ) //&& (wrkList[pointer].OnDisplay = true))
                {
                    XPos = XPos + wrkList[firstStringPointer].TextWidth;
                    wrkList[firstStringPointer].OnDisplay = false;

                    firstStringPointer ++;
                    if (firstStringPointer == wrkList.Count)
                    { firstStringPointer = 0; }
                }
                counter++;
            }

            tickToTickEnd = sw1.ElapsedMilliseconds;

            if (tickToTickEnd > maxt) { maxt = tickToTickEnd; }
            if (tickToTickEnd < mint) { mint = tickToTickEnd; }
            sleeptime = (35 - (int)tickToTickEnd); //35
            if (sleeptime > 0) { Thread.Sleep(sleeptime); }

            sw1.Stop();
            tot = sw1.ElapsedMilliseconds;
            sw1.Reset();
            sw1.Start();
            parent.Invalidate ();
            //parent.Refresh();
            
        }

        private void DoubleBuffering_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
        {
            tickendToPaint = sw1.ElapsedMilliseconds;
            Graphics grfx = e.Graphics;
            grfx.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            //grfx.Flush(System.Drawing.Drawing2D.FlushIntention.Flush);
           // grfx.Flush(System.Drawing.Drawing2D.FlushIntention.Sync);
            //System.Drawing.Drawing2D.FlushIntention.Flush ; g;
            //Pen drawingPen = new Pen(Color.Red, 10);
      
            int c = firstStringPointer;
            int nextPos = XPos;

            while (wrkList[c].OnDisplay && (nextPos<parent.Width))
            {
                if (wrkList[c].TextBackColor!= null)
                {
                    grfx.FillRectangle(wrkList[c].TextBackColor, nextPos, 0, wrkList[c].TextWidth, wrkList[c].TextHeight);
                }
                
                grfx.DrawString(wrkList[c].TextString, wrkList[c].TextFont, new SolidBrush(wrkList[c].TextColor), nextPos, YPos);
                nextPos=nextPos+wrkList[c].TextWidth;
                c++;
                if (c >= wrkList.Count)
                { c = 0; }
            }

            paintToEndPaint = sw1.ElapsedMilliseconds;
        }


        #region Start ,Stop ,Reset Srolling
        public void StartScrolling()
        { 
            if (wrkList.Count == 0) {return;}
            parent.Paint += new System.Windows.Forms.PaintEventHandler(DoubleBuffering_Paint);
            XPos = ((rightToLeft == true) ?  parent.Width : XPos );
            wrkList[0].OnDisplay = true;
            dispTimer.Start();
            scrollRunning = true;
        }


        public void StopScrolling()
        {
            parent.Paint -= new System.Windows.Forms.PaintEventHandler(this.DoubleBuffering_Paint);
            dispTimer.Stop ();
            scrollRunning = false;
        }


        public void Reset()
        {
            XPos = parent.Width;
            if (wrkList.Count != 0) { wrkList[0].OnDisplay = true; ; }
        }


    }
    #endregion
#endregion

}
