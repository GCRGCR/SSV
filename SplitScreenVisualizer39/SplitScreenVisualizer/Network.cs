using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;
using System.Threading;
using System.Drawing;

namespace SplitScreenVisualizer
{
    class Network
    {
        public static event EventHandler ReceiveMessageEvent;
        public static bool ReciveEnabled;
        static ManualResetEvent done = new ManualResetEvent(false);

        public struct receiveData
        {
            public string ID;
            public Color bg1Color;
            public Color bg2Color;
            public byte byteTimeOver;
            public DateTime EndDate1;
            public DateTime EndDate2;
            public string strText1;
            public string strText2;
            public bool bRemoteSet;
            public bool Valid;
            public bool bLineMessage;

            public uint showOnPrimary;
            public string strLineMessage;
            public int intLineMessageTime;
            public int intLineMessageFS;

            public int PWidth;
            public int rxPort;
            public string strPPath;
            public int iFS;
        }
        public static receiveData Message = new receiveData();

        static object locker = new object();

        public static void  ServerStart(object Parameter)
        {
            lock (locker)
            {
                Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name + "--" + Thread.CurrentThread.Name); //---------
                ReciveEnabled = true;
                /* Socket client = null;
            
                 IPAddress[] localIPs = Dns.GetHostAddresses(Dns.GetHostName());
                 IPEndPoint localEndPoint = new IPEndPoint(localIPs[2], Convert.ToInt32 (Parameter));
                 Socket listenSocket = new Socket(localEndPoint.AddressFamily, SocketType.Dgram, ProtocolType.Udp );
                 listenSocket.Bind(localEndPoint);
                 listenSocket.Listen(2);
                 client = listenSocket.Accept();

                 Debug.WriteLine("Kliens bejelentkezett, a következő IP címről: {0}", ((IPEndPoint)client.RemoteEndPoint).Address.ToString());
                 byte[] data = new byte[256];
                 int length = client.Receive(data);
                 Debug.WriteLine("A kliens üzenete: {0}",Encoding.ASCII.GetString(data, 0, length));*/

                /*         //Creates a UdpClient for reading incoming data.
                         UdpClient receivingUdpClient = new UdpClient(11000);

                         //Creates an IPEndPoint to record the IP Address and port number of the sender.  // The IPEndPoint will allow you to read datagrams sent from any source.
                         IPEndPoint RemoteIpEndPoint = new IPEndPoint(IPAddress.Any, 0);
                         try
                         {
                             reciveData data =new reciveData();
                             // Blocks until a message returns on this socket from a remote host.
                             Byte[] receiveData = receivingUdpClient.Receive(ref RemoteIpEndPoint);
                             Debug.WriteLine(reciveData.
                             string returnData = Encoding.ASCII.GetString(receiveData);
                
                             Debug.WriteLine("This is the message you received " + returnData.ToString());
                             Debug.WriteLine("This message was sent from " + RemoteIpEndPoint.Address.ToString() +
                                              " on their port number " + RemoteIpEndPoint.Port.ToString());

                         }
                         catch (Exception e)
                         {
                             Debug.WriteLine(e.ToString());
                         }
 
                    */
                Debug.WriteLine(Convert.ToString(Thread.CurrentThread.ManagedThreadId), "Szál-Id: {0}");
                int listenPort = Convert.ToInt32(Parameter);

                try
                {
                    IPEndPoint endpoint = new IPEndPoint(IPAddress.Any, listenPort);
                    UdpClient client = new UdpClient(endpoint);

                    UdpState oState = new UdpState();
                    oState.ep = endpoint;
                    oState.cl = client;

                    while (ReciveEnabled)
                    {
                        Debug.WriteLine("listening for messages");
                        done.Reset();
                        client.BeginReceive(new AsyncCallback(ReceiveCallback), oState);
                        // Do some work while we wait for a message. For this example, // we'll just sleep 
                        done.WaitOne();
                    }
                    //client.EndReceive( IAsyncResult ar, ref endpoint);
                    //ReceiveCallback(null);
                    client.Client.Shutdown(SocketShutdown.Both);
                    client.Client.Close();
                    client.Close();
                    client = null;
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show(e.Message.ToString());
                }
            }
        }

        public static void ServerStop()
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            ReciveEnabled = false;
            done.Set();
        }

        public static void ReceiveCallback(IAsyncResult ar)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            
            UdpClient cli = (UdpClient)((UdpState)(ar.AsyncState)).cl;
            IPEndPoint endp = (IPEndPoint)((UdpState)(ar.AsyncState)).ep;
            done.Set();
            try
            {
                Byte[] receiveBytes = cli.EndReceive(ar, ref endp);
                string receiveString = Encoding.Default.GetString(receiveBytes);

                Debug.WriteLine( receiveString,"Received: {0}");

                if (StoreMessage(receiveBytes))
                {
                    byte[] response = new byte[2];
                    response = Encoding.ASCII.GetBytes("SSV");
                    cli.Send(response, response.Length, endp);
                    if (ReceiveMessageEvent != null)
                        ReceiveMessageEvent(Message, new EventArgs());
                }
            }
            catch (System.ObjectDisposedException)
            {
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.ToString());
            }
        }

        public static bool StoreMessage(byte[] recivedBytes)
        {
            Debug.Print(System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name); //---------
            Message.ID = Encoding.Default.GetString(recivedBytes, 0, 3);

            if (Message.ID == "SSV")
            {
                Message.Valid = true;
                Message.bg1Color = Color.FromArgb(recivedBytes[3], recivedBytes[4], recivedBytes[5]);
                Message.bg2Color = Color.FromArgb(recivedBytes[6], recivedBytes[7], recivedBytes[8]);
                Message.byteTimeOver = recivedBytes[9];
                string msgText = Encoding.Default.GetString(recivedBytes).Remove(0, 100);
                if (0 < msgText.IndexOf(char.ConvertFromUtf32(13)))
                {
                    Message.strText1 = msgText.Substring(0, msgText.IndexOf(char.ConvertFromUtf32(13)));
                    Message.strText2 = msgText.Substring(100,msgText.IndexOf(char.ConvertFromUtf32(13),100)-100);
                    if (1 < Message.strText1.Length) 
                    {                        
                        string datestring= Encoding.Default.GetString(recivedBytes,10, 20);
                        Message.EndDate1 = Convert.ToDateTime(datestring);
                    }

                    if (1 < Message.strText2.Length) 
                    {
                        string datestring = Encoding.Default.GetString(recivedBytes, 30, 20);
                        Message.EndDate2 = Convert.ToDateTime(datestring);
                    }

                    Debug.WriteLine(Message.strText1);
                    Debug.WriteLine(Message.strText2);
                }
                
            }
            else if (Message.ID == "SSW")
            {
                Message.bRemoteSet = true;
                Message.PWidth= recivedBytes[3] * 256 + recivedBytes[4];
                Message.rxPort = recivedBytes[5] * 256 + recivedBytes[6];
                Message.iFS=recivedBytes[7] * 256 + recivedBytes[8];
                string msgText = Encoding.Default.GetString(recivedBytes).Remove(0, 19);
                Message.strPPath = msgText.Substring(0, msgText.IndexOf(char.ConvertFromUtf32(13)));
            }
            else if (Message.ID == "SSM")
            {
                Message.strLineMessage = Encoding.Default.GetString(recivedBytes);
                if (Message.strLineMessage.Length ==40)
                {
                    Message.bLineMessage = true;
                    Message.strLineMessage = Message.strLineMessage.Remove( 0, 4);
                    Message.strLineMessage = Message.strLineMessage.Remove(32, 4);
                    Message.strLineMessage = Message.strLineMessage.Trim();
                    
                    if ((recivedBytes[3] & (byte)1) == 1) { Message.showOnPrimary = 1; }
                    else { Message.showOnPrimary = 0;}
                    
                    Message.intLineMessageTime = recivedBytes[36] + recivedBytes[37] * 256;
                    Message.intLineMessageFS = recivedBytes[38]  + recivedBytes[39] * 256;
                } 
            }

            else { return false; }
            return true;
        }

        public class UdpState
        {
            public IPEndPoint ep;
            public UdpClient cl;
        }

    }
}
