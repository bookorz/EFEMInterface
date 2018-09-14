using EFEMInterface.MessageInterface;
using log4net;
using SANWA.Utility.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EFEMInterface.Comm
{
    // State object for reading client data asynchronously  
    public class StateObject
    {
        // Client  socket.  
        public Socket workSocket = null;
        // Size of receive buffer.  
        public const int BufferSize = 1024;
        // Receive buffer.  
        public byte[] buffer = new byte[BufferSize];
        // Received data string.  
        public StringBuilder sb = new StringBuilder();

        //public void ClearBuffer()
        //{
        //    buffer = new byte[BufferSize];
        //}
    }

    public class SocketServer
    {
        ICommMessage _EventReport;
        // Thread signal.  
        public ManualResetEvent allDone = new ManualResetEvent(false);

        public SocketServer(ICommMessage EventReport)
        {
            _EventReport = EventReport;
            ThreadPool.QueueUserWorkItem(new WaitCallback(StartListening));

        }

        public void StartListening(object msg)
        {
            // Establish the local endpoint for the socket.  
            // The DNS name of the computer  

            //IPHostEntry ipHostInfo = Dns.GetHostEntry(Dns.GetHostName());
            //IPAddress ipAddress = IPAddress.Parse("127.0.0.1");
            IPAddress ipAddress = IPAddress.Parse(SystemConfig.Get().EFEMInterfaceConn);
            IPEndPoint localEndPoint = new IPEndPoint(ipAddress, 13000);

            // Create a TCP/IP socket.  
            Socket listener = new Socket(ipAddress.AddressFamily,
                SocketType.Stream, ProtocolType.Tcp);

            // Bind the socket to the local endpoint and listen for incoming connections.  
            try
            {
                listener.Bind(localEndPoint);
                listener.Listen(1);
                _EventReport.On_Connection_Connecting();
                while (true)
                {
                    // Set the event to nonsignaled state.  
                    allDone.Reset();

                    // Start an asynchronous socket to listen for connections.  
                    //Console.WriteLine("Waiting for a connection...");
                    listener.BeginAccept(
                        new AsyncCallback(AcceptCallback),
                        listener);

                    // Wait until a connection is made before continuing.  
                    allDone.WaitOne();
                }

            }
            catch (Exception e)
            {
                //Console.WriteLine(e.ToString());
            }

            //Console.WriteLine("\nPress ENTER to continue...");
            //Console.Read();

        }

        public void AcceptCallback(IAsyncResult ar)
        {
            // Signal the main thread to continue.  
            allDone.Set();

            // Get the socket that handles the client request.  
            Socket listener = (Socket)ar.AsyncState;
            Socket handler = listener.EndAccept(ar);
            _EventReport.On_Connection_Connected(handler);
            // Create the state object.  
            StateObject state = new StateObject();
            state.workSocket = handler;
            handler.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0,
                new AsyncCallback(ReadCallback), state);
        }

        public void ReadCallback(IAsyncResult ar)
        {
            try
            {
                String content = String.Empty;

                // Retrieve the state object and the handler socket  
                // from the asynchronous state object.  
                StateObject state = (StateObject)ar.AsyncState;
                Socket handler = state.workSocket;

                // Read data from the client socket.   

                int bytesRead = handler.EndReceive(ar);

                if (bytesRead > 0)
                {
                    // There  might be more data, so store the data received so far.  
                    state.sb.Append(Encoding.ASCII.GetString(
                        state.buffer, 0, bytesRead));

                    // Check for end-of-file tag. If it is not there, read   
                    // more data.  
                    content = state.sb.ToString();
                    if (content.IndexOf("\r") > -1)
                    {
                        // All the data has been read from the   
                        // client. Display it on the console.  
                        //Console.WriteLine("Read {0} bytes from socket. \n Data : {1}",
                        //    content.Length, content);
                        state.sb.Clear();
                        _EventReport.On_Connection_Message(handler, content);
                        //state.ClearBuffer();
                        handler.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0,
                       new AsyncCallback(ReadCallback), state);
                    }
                    else
                    {
                        // Not all data received. Get more.  
                        handler.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0,
                        new AsyncCallback(ReadCallback), state);
                    }
                }
            }
            catch (Exception e)
            {
                _EventReport.On_Connection_Disconnected();
            }
        }

        public void Send(Socket handler, String data)
        {
            // Convert the string data to byte data using ASCII encoding.  
            byte[] byteData = Encoding.ASCII.GetBytes(data);

            // Begin sending the data to the remote device.  
            try
            {
                handler.BeginSend(byteData, 0, byteData.Length, 0,
                new AsyncCallback(SendCallback), handler);
            }
            catch (Exception e)
            {
                _EventReport.On_Connection_Disconnected();
            }
        }

        private void SendCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the socket from the state object.  
                Socket handler = (Socket)ar.AsyncState;

                // Complete sending the data to the remote device.  
                int bytesSent = handler.EndSend(ar);
                //Console.WriteLine("Sent {0} bytes to client.", bytesSent);

                //handler.Shutdown(SocketShutdown.Both);
                //handler.Close();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
