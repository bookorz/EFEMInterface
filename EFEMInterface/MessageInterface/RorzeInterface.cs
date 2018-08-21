using EFEMInterface.Comm;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace EFEMInterface.MessageInterface
{
    public class RorzeInterface : ICommMessage, IHandlingTimeOutReport
    {
        ILog logger = LogManager.GetLogger(typeof(RorzeInterface));

        private Dictionary<string, OnHandling> OnHandlingCmds = new Dictionary<string, OnHandling>();

        IEFEMControl _EventReport;

        SocketServer Comm;

        public RorzeInterface(IEFEMControl EventReport)
        {
            _EventReport = EventReport;
            Comm = new SocketServer(this);
        }




        public class RorzeCommand
        {
            public string OrgMsg = "";
            public string Command;
            public List<string> Parameter = new List<string>();
            public string CommandType = "";
        }

        public class CommandType
        {
            public const string MOV = "MOV";
            public const string SET = "SET";
            public const string GET = "GET";
            public const string ACK = "ACK";
            public const string ABS = "ABS";
            public const string INF = "INF";
        }

        public class TransactionType
        {
            public const string Excuted = "Excuted";
            public const string Finished = "Finished";
            public const string Error = "Error";
            public const string Event = "Event";
            public const string Information = "Information";
            public const string ReInformation = "ReInformation";
            public const string Abnormal = "Abnormal";
        }

        private RorzeCommand CmdParser(string Msg)
        {
            RorzeCommand each = new RorzeCommand();

            each.OrgMsg = Msg;

            string[] content = Msg.Replace(";", "").Replace("\r", "").Split(':', '/');

            for (int i = 0; i < content.Length; i++)
            {
                switch (i)
                {
                    case 0:

                        each.CommandType = content[i];
                        break;
                    case 1:

                        each.Command = content[i];

                        break;
                    default:
                        each.Parameter.Add(content[i]);
                        break;
                }
            }


            return each;
        }

        private string CmdAssembler(RorzeCommand cmd, string CommandType)
        {
            string result = "";

            result = CommandType + ":" + cmd.Command;

            foreach (string param in cmd.Parameter)
            {
                result += "/" + param;
            }

            result += ";\r";

            return result;
        }

        private string ErrorAssembler(RorzeCommand cmd, string Param1, string Param2)
        {
            string result = "";

            result = "ABS:" + cmd.Command;

            foreach (string param in cmd.Parameter)
            {
                result += "/" + param;
            }
            result += "|ERROR";

            if (!Param1.Equals(""))
            {
                result += "/" + Param1;
            }
            if (!Param2.Equals(""))
            {
                result += "/" + Param2;
            }
            result += ";\r";

            return result;
        }

        private string CancelAssembler(RorzeCommand cmd, string Factor, string Place)
        {
            string result = "";

            result = "CAN:" + cmd.Command;

            foreach (string param in cmd.Parameter)
            {
                result += "/" + param;
            }
            if (!Factor.Equals(""))
            {
                result += "|" + Factor;
            }
            if (!Factor.Equals(""))
            {
                result += "/" + Place;
            }
            result += ";\r";
            return result;
        }

        public void On_Connection_Connecting()
        {
            _EventReport.On_Connection_Connecting();
        }

        public void On_Connection_Connected(Socket handler)
        {
            RorzeCommand CommunityActive = new RorzeCommand();
            CommunityActive.CommandType = CommandType.INF;
            CommunityActive.Command = "READY";
            CommunityActive.Parameter.Add("COMM*");
            string CommandMsg = CmdAssembler(CommunityActive, CommandType.INF);
            Comm.Send(handler, CommandMsg);
            _EventReport.On_Connection_Connected();
        }

        public void On_Connection_Disconnected()
        {
            _EventReport.On_Connection_Disconnected();
        }

        public void On_Connection_Error(string Msg)
        {
            _EventReport.On_CommandMessage(Msg.ToString());
        }

        public void On_Connection_Message(Socket handler, string content)
        {
            try
            {
                _EventReport.On_CommandMessage("Recv:" + content.ToString());
                RorzeCommand cmd = CmdParser(content);
                string key = cmd.Command;
               

                switch (cmd.CommandType)
                {
                    case CommandType.GET:
                    case CommandType.SET:
                        //回報收到訊息
                        string CommandMsg = CmdAssembler(cmd, CommandType.ACK);
                        Comm.Send(handler, CommandMsg);
                        _EventReport.On_CommandMessage("Send:" + CommandMsg);
                        break;
                    case CommandType.MOV:
                        //回報收到訊息
                        CommandMsg = CmdAssembler(cmd, CommandType.ACK);
                        Comm.Send(handler, CommandMsg);
                        _EventReport.On_CommandMessage("Send:" + CommandMsg);
                        if (OnHandlingCmds.ContainsKey(key))
                        {//已有相同指令正在執行中，回覆錯誤訊息給上位系統

                            string ErrorMessage = CancelAssembler(cmd, ErrorCategory.CancelFactor.BUSY, ErrorCategory.CancelPlace.DUPLICATE);
                            Comm.Send(handler, ErrorMessage);
                            _EventReport.On_CommandMessage("Send:" + ErrorMessage);
                        }
                        else
                        {
                            OnHandling WaitForHandle = new OnHandling(this);
                            WaitForHandle.Cmd = cmd;
                            WaitForHandle.Handler = handler;
                            OnHandlingCmds.Add(key, WaitForHandle);

                            //處理邏輯開始


                            //處理邏輯結束
                            CommandMsg = CmdAssembler(cmd, CommandType.INF);//傳送動作完成給上位系統
                            Comm.Send(handler, CommandMsg);
                            _EventReport.On_CommandMessage("Send:" + CommandMsg);
                            WaitForHandle.SetTimeOutMonitor(true);//設定Timeout監控開始，5秒後
                        }
                        break;
                    case CommandType.ACK://收到上位系統回覆
                        if (OnHandlingCmds.ContainsKey(key))//確認有這筆
                        {
                            OnHandling WaitForHandle;
                            OnHandlingCmds.TryGetValue(key, out WaitForHandle);
                            WaitForHandle.SetTimeOutMonitor(false);//設定Timeout監控停止
                            OnHandlingCmds.Remove(key);//從待處理名單移除
                        }
                        break;
                }




            }
            catch (Exception e)
            {
                _EventReport.On_CommandMessage(e.StackTrace);
            }
        }

        public void On_Handling_TimeOut(OnHandling TimeOutCmd)
        {
            string key = TimeOutCmd.Cmd.Command;
            if (TimeOutCmd.INF_RetryCount < 3)
            {
                TimeOutCmd.INF_RetryCount++;
                string CommandMsg = CmdAssembler(TimeOutCmd.Cmd, CommandType.INF);
                Comm.Send(TimeOutCmd.Handler, CommandMsg);
                _EventReport.On_CommandMessage("Send:" + CommandMsg);
            }
            else
            {
                TimeOutCmd.SetTimeOutMonitor(false);//設定Timeout監控停止
                OnHandlingCmds.Remove(key);//從待處理名單移除
                _EventReport.On_CommandMessage("Retry Timeout:" + CmdAssembler(TimeOutCmd.Cmd, CommandType.INF));
            }
        }


    }
}
