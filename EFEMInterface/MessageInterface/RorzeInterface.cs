using EFEMInterface.Comm;
using log4net;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using TransferControl.Engine;
using TransferControl.Management;

namespace EFEMInterface.MessageInterface
{
    public class RorzeInterface : ICommMessage, IHandlingTimeOutReport, IHostInterfaceReport
    {
        ILog logger = LogManager.GetLogger(typeof(RorzeInterface));

        private ConcurrentDictionary<string, OnHandling> OnHandlingCmds = new ConcurrentDictionary<string, OnHandling>();

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
            public const string NAK = "NAK";
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

            string[] content = Msg.Replace(";", "").Replace("\r", "").Split(':', '/', '>');

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

        private string CmdFormatErrorAssembler(RorzeCommand cmd)
        {
            string result = "";

            result = "NAK:MSG|" + cmd.Command;

            foreach (string param in cmd.Parameter)
            {
                result += "/" + param;
            }

            result += ";\r";

            return result;
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
            if (!Place.Equals(""))
            {
                result += "/" + Place;
            }
            result += ";\r";
            return result;
        }

        private string InfoAssembler(RorzeCommand cmd, string data1, string data2, bool isEvt = false)
        {
            string result = "";
            if (isEvt)
            {
                result = "EVT:" + cmd.Command;
            }
            else
            {
                result = "INF:" + cmd.Command;
            }

            switch (cmd.Command)
            {
                case "MAPDT":
                case "TRANSREQ":
                case "CLAMP":
                case "STATE":
                case "MODE":
                case "EVENT":
                case "CSTID":
                case "SIZE":
                    //result += "/" + cmd.Parameter[0] + "/" + data;
                    result += "/" + cmd.Parameter[0];
                    if (!data1.Equals(""))
                    {
                        result += "/" + data1;
                    }
                    break;

                case "SIGSTAT":
                    //result += "/" + cmd.Parameter[0] + "/" + data;
                    result += "/" + cmd.Parameter[0];
                    if (!data1.Equals(""))
                    {
                        result += "/" + data1;
                    }
                    if (!data2.Equals(""))
                    {
                        result += "/" + data2;
                    }
                    break;
                case "ERROR":
                    if (!data1.Equals(""))
                    {
                        result += "/" + data1;
                    }
                    if (!data2.Equals(""))
                    {
                        result += "/" + data2;
                    }
                    break;
            }
            result += ";\r";
            return result;
        }

        private void SendAck(OnHandling WaitForHandle)
        {
            //回報收到訊息
            string CommandMsg = CmdAssembler(WaitForHandle.Cmd, CommandType.ACK);

            Comm.Send(WaitForHandle.Handler, CommandMsg);
            _EventReport.On_CommandMessage("Send:" + CommandMsg);
        }

        private void SendNak(OnHandling WaitForHandle, string detail)
        {
            string ErrorMsg = CmdFormatErrorAssembler(WaitForHandle.Cmd);
            Comm.Send(WaitForHandle.Handler, ErrorMsg);
            _EventReport.On_CommandMessage("Err :" + detail);
            _EventReport.On_CommandMessage("Send:" + ErrorMsg);
        }

        private void SendCancel(OnHandling WaitForHandle, string Factor, string Place, string detail)
        {
            //回報設備不可使用
            if (!WaitForHandle.IsReturn)
            {
                //WaitForHandle.IsReturn = true;
                string CancelMsg = CancelAssembler(WaitForHandle.Cmd, Factor, Place);
                Comm.Send(WaitForHandle.Handler, CancelMsg);
                _EventReport.On_CommandMessage("Err :" + detail);
                _EventReport.On_CommandMessage("Send:" + CancelMsg);
            }
        }

        private void SendInfo(OnHandling WaitForHandle)
        {
            if (!WaitForHandle.IsReturn)
            {
                WaitForHandle.IsReturn = true;
                string CommandMsg = CmdAssembler(WaitForHandle.Cmd, CommandType.INF);//傳送動作完成給上位系統
                WaitForHandle.NotConfirmMsg = CommandMsg;
                WaitForHandle.SetTimeOutMonitor(true);//設定Timeout監控開始，5秒後
                Comm.Send(WaitForHandle.Handler, CommandMsg);
                _EventReport.On_CommandMessage("Send:" + CommandMsg);

            }
        }

        private void SendInfo(OnHandling WaitForHandle, string data1, string data2)
        {
            if (!WaitForHandle.IsReturn)
            {
                WaitForHandle.IsReturn = true;
                string CommandMsg = InfoAssembler(WaitForHandle.Cmd, data1, data2);//回傳資料給上位系統
                WaitForHandle.NotConfirmMsg = CommandMsg;
                WaitForHandle.SetTimeOutMonitor(true);//設定Timeout監控開始，5秒後
                Comm.Send(WaitForHandle.Handler, CommandMsg);
                _EventReport.On_CommandMessage("Send:" + CommandMsg);

            }
        }

        private void SendABS(OnHandling WaitForHandle, string ErrorType)
        {
            if (!WaitForHandle.IsReturn)
            {
                WaitForHandle.IsReturn = true;
                string param1 = "";
                string param2 = "";
                switch (ErrorType)
                {
                    default:
                        param1 = ErrorType;
                        param2 = "TEST";
                        break;
                }

                string ErrorMsg = ErrorAssembler(WaitForHandle.Cmd, param1, param2);
                WaitForHandle.SetTimeOutMonitor(true);//設定Timeout監控開始，5秒後
                Comm.Send(WaitForHandle.Handler, ErrorMsg);
                WaitForHandle.NotConfirmMsg = ErrorMsg;
                // _EventReport.On_CommandMessage("Err :" + detail);
                _EventReport.On_CommandMessage("Send:" + ErrorMsg);

            }
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
            CommunityActive.Parameter.Add("COMM");

            OnHandling WaitForHandle = new OnHandling(this);
            WaitForHandle.Cmd = CommunityActive;
            WaitForHandle.Handler = handler;
            OnHandlingCmds.TryAdd(WaitForHandle.ID, WaitForHandle);

            SendInfo(WaitForHandle);

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
                int no = 0;
                _EventReport.On_CommandMessage("Recv:" + content.ToString());
                RorzeCommand cmd = CmdParser(content);

                OnHandling WaitForHandle = null;
                switch (cmd.CommandType)
                {
                    case CommandType.GET:
                    case CommandType.SET:
                    case CommandType.MOV:
                        WaitForHandle = new OnHandling(this);
                        WaitForHandle.Cmd = cmd;
                        WaitForHandle.Handler = handler;


                        var findHandling = from Handling in OnHandlingCmds.Values.ToList()
                                           where Handling.Cmd.Command.Equals(cmd.Command) && Handling.Cmd.CommandType.Equals(CommandType.MOV)

                                           select Handling;
                        OnHandlingCmds.TryAdd(WaitForHandle.ID, WaitForHandle);
                        if (findHandling.Count() != 0)
                        {

                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.BUSY, "DUPLICATE", "Command already exsit.");
                            SendInfo(WaitForHandle);
                            return;
                        }


                        break;
                    case CommandType.ACK://收到上位系統回覆
                        List<OnHandling> tmp = OnHandlingCmds.Values.ToList();
                        tmp.Sort((x, y) => { return x.ReceiveTime.CompareTo(y.ReceiveTime); });

                        findHandling = from Handling in tmp
                                       where Handling.Cmd.Command.Equals(cmd.Command)
                                       select Handling;

                        if (findHandling.Count() != 0)
                        {
                            WaitForHandle = findHandling.First();
                            WaitForHandle.SetTimeOutMonitor(false);//設定Timeout監控停止
                            OnHandlingCmds.TryRemove(WaitForHandle.ID, out WaitForHandle);//從待處理名單移除
                        }


                        break;
                    default:
                        //命令錯誤
                        SendNak(WaitForHandle, "Command format error.");
                        SendInfo(WaitForHandle);
                        return;
                }

                //處理邏輯開始
                Node node = null;
                Transaction txn;
                switch (cmd.CommandType)
                {
                    #region GET

                    case CommandType.GET:
                        switch (cmd.Command.ToUpper())
                        {
                            case "MAPDT":
                                //取得LoadPort Mapping 結果
                                #region MAPDT
                                try
                                {

                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    node = NodeManagement.Get(NodeNameConvert(cmd.Parameter[0], "LOADPORT"));

                                    if (node != null)
                                    {
                                        SendAck(WaitForHandle);//發送ACK給上位系統
                                        txn = new Transaction();
                                        txn.FormName = WaitForHandle.ID;
                                        txn.Method = Transaction.Command.LoadPortType.GetMapping;
                                        //node.SendCommand(txn);//下GetMapping命令給Loadport


                                        //*********************test begin*********************

                                        SANWA.Utility.ReturnMessage Msg = new SANWA.Utility.ReturnMessage();
                                        Msg.Command = Transaction.Command.LoadPortType.GetMapping;
                                        Msg.CommandType = "GET";
                                        Msg.Value = "1111111111111111111111111";


                                        //*********************test   end*********************
                                    }
                                    else
                                    {
                                        //回報設備不可使用
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "Loadport not found.");
                                        SendInfo(WaitForHandle);
                                    }

                                }
                                catch
                                {
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                    break;
                                }
                                #endregion
                                break;
                            case "ERROR":
                                //取得ERROR狀態
                                //檢查命令格式
                                try
                                {
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    //*********************test begin*********************
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle, "COMMAND", "ROBOT");

                                    //*********************test   end*********************
                                }
                                catch
                                {
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                    break;
                                }
                                break;
                            case "CLAMP":
                                //取得CLAMP狀態
                                try
                                {
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("ARM") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ARM", ""), out no) &&
                                                    cmd.Parameter[i].Replace("ARM", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                   int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                   cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle, "ON", "");


                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "STATE":
                                //取得STATE狀態

                                //*********************test begin*********************
                                try
                                {
                                    string returnValue = "";
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                if (cmd.Parameter[i].Equals("VER"))
                                                {
                                                    returnValue = "1.0.0.1(2018-08-01)";
                                                }
                                                else if (cmd.Parameter[i].Equals("TRACK"))
                                                {
                                                    returnValue = "NONE/200/300";
                                                }
                                                else if (cmd.Parameter[i].IndexOf("PRS") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("PRS", ""), out no) &&
                                                    cmd.Parameter[i].Replace("PRS", "").Length == 1)
                                                {
                                                    returnValue = "SNO1|00000000,SNO2|00003000,SNO3|00009527,SNO4|88888888";
                                                }
                                                else if (cmd.Parameter[i].IndexOf("FFU") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("FFU", ""), out no) &&
                                                    cmd.Parameter[i].Replace("FFU", "").Length == 1)
                                                {
                                                    returnValue = "FNO1|00000000,FNO2|00003000,FNO3|00009527,FNO4|88888888";
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle, returnValue, "");

                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                //*********************test   end*********************

                                break;
                            case "MODE":
                                //取得MODE狀態

                                //*********************test begin*********************
                                try
                                {
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle, "MANUAL", "");
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                //*********************test   end*********************

                                break;
                            case "TRANSREQ":
                                //取得E84狀態
                                try
                                {
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle, "STOP", "");
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                //*********************test begin*********************

                                //*********************test   end*********************

                                break;
                            case "SIGSTAT":
                                //取得SIGSTAT狀態
                                try
                                {
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("SYSTEM"))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle, "11111111111111111111111111111111", "11111111111111111111111111111111");
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                //*********************test begin*********************


                                //*********************test   end*********************

                                break;
                            case "EVENT":
                                //取得EVENT狀態

                                //*********************test begin*********************
                                try
                                {
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].Equals("MAPDT"))
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("TRANSREQ"))
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("SYSTEM"))
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("PORT"))
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("PRS"))
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("FFU"))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle, "OFF", "");

                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                //*********************test   end*********************

                                break;
                            case "CSTID":
                                //取得CSTID

                                //*********************test begin*********************
                                try
                                {
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle, "FOUPIDXX", "");
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                //*********************test   end*********************

                                break;
                            case "SIZE":
                                //取得Wafer Size

                                //*********************test begin*********************
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                //Parameter 1 designates transfer source, and Parameter 4 designates transfer destination.
                                                if (cmd.Parameter[i].IndexOf("ARM") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ARM", ""), out no) &&
                                                    cmd.Parameter[i].Replace("ARM", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 3)
                                                {

                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                    cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {

                                                }
                                                else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("LL", ""), out no) &&
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 3))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:

                                                //Parameter 2 designates the End-EF used for carrying out. Parameter 3 designates the End-EF used for carrying in.
                                                if ((cmd.Parameter[i].Equals("200") || cmd.Parameter[i].Equals("300")))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle, "300", "");

                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                //*********************test   end*********************

                                break;
                            default:
                                //命令錯誤
                                SendNak(WaitForHandle, "Command format error.");
                                SendInfo(WaitForHandle);
                                break;
                        }
                        break;
                    #endregion
                    #region SET
                    case CommandType.SET:
                        switch (cmd.Command.ToUpper())
                        {
                            case "ALIGN":
                                //設定Aligner旋轉Notch角度
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                   int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                   cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                  int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                  cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                 int.TryParse(cmd.Parameter[i].Replace("LL", ""), out no) &&
                                                 cmd.Parameter[i].Replace("LL", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].IndexOf("D") != -1 &&
                                                 int.Parse(cmd.Parameter[i].Replace("D", "")) < 180000 &&
                                                 int.Parse(cmd.Parameter[i].Replace("D", "")) > -180000)
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    node = NodeManagement.Get(NodeNameConvert(cmd.Parameter[0], "ALIGNER"));//取得Aligner

                                    if (node != null)
                                    {
                                        //**************************取得補正角度值**************************Begin
                                        string DestName = "";
                                        if (cmd.Parameter[1].IndexOf("P") != -1)//指定UnloadPort 
                                        {
                                            DestName = NodeNameConvert(cmd.Parameter[1], "LOADPORT");
                                        }
                                        else if (cmd.Parameter[1].IndexOf("LL") != -1)//指定Load Lock stage 
                                        {
                                            DestName = NodeNameConvert(cmd.Parameter[1], "STAGE");
                                        }

                                        Node dest = NodeManagement.Get(DestName);//取得目的地UnloadPort
                                        Node NextRobot = NodeManagement.GetNextRobot(DestName);//取得搬送Robot
                                        if (dest != null)
                                        {
                                            if (NextRobot != null)
                                            {

                                                RobotPoint ptAligner = PointManagement.GetPoint(NextRobot.Name, node.Name, dest.WaferSize);
                                                if (ptAligner != null)
                                                {
                                                    int Offset = ptAligner.Offset;//Aligner補償值


                                                    RobotPoint ptDest = PointManagement.GetPoint(NextRobot.Name, dest.Name, dest.WaferSize);
                                                    if (ptDest != null)
                                                    {
                                                        Offset += ptDest.Offset;//加上目的地補償角度
                                                        Offset += dest.NotchAngle;//加上Notch角度位置
                                                        node.DesignatesAngle = Offset.ToString();//事先存入Aligner屬性中
                                                        SendAck(WaitForHandle);
                                                        SendInfo(WaitForHandle);

                                                    }
                                                    else
                                                    {
                                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "ptDest not found.");
                                                        SendInfo(WaitForHandle);
                                                    }

                                                }
                                                else
                                                {
                                                    SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "ptAligner not found.");
                                                    SendInfo(WaitForHandle);
                                                }
                                            }
                                            else
                                            {
                                                SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "NextRobot not found.");
                                                SendInfo(WaitForHandle);
                                            }
                                        }
                                        else if (cmd.Parameter[1].IndexOf("D") != -1)//指定角度
                                        {
                                            string angle = cmd.Parameter[1].Replace("D", "");
                                            if (angle.Length >= 6)
                                            {
                                                angle = angle.Substring(angle.Length - 6, 3);
                                                node.DesignatesAngle = angle;//事先存入Aligner屬性中

                                                SendAck(WaitForHandle);
                                            }
                                            else
                                            {
                                                SendNak(WaitForHandle, "Command format error.");
                                            }
                                            SendInfo(WaitForHandle);

                                        }
                                        else
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "NextRobot not found.");
                                            SendInfo(WaitForHandle);
                                        }
                                    }
                                    else
                                    {
                                        //回報設備不可使用

                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "Aligner not found.");
                                        SendInfo(WaitForHandle);
                                    }

                                }
                                catch
                                {
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                    break;
                                }
                                break;
                            case "ERROR":
                                try
                                {
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                if (cmd.Parameter[i].Equals("CLEAR"))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }
                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    string TaskName = "SET_ERROR";
                                    //Dictionary<string, string> param = new Dictionary<string, string>();


                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "CLAMP":
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Target = "";
                                    string Arm = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                if (cmd.Parameter[i].IndexOf("ARM") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ARM", ""), out no) &&
                                                    cmd.Parameter[i].Replace("ARM", "").Length == 1)
                                                {
                                                    Target = "ROBOT01";
                                                    Arm = no.ToString();
                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                   int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                   cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "ALIGNER");
                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    Target = NodeNameConvert("ALIGN1", "ALIGNER");
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:
                                                if (cmd.Parameter[i].Equals("ON") )
                                                {
                                                    TaskName = "SET_CLAMP_ON";
                                                }
                                                else if (cmd.Parameter[i].Equals("OFF"))
                                                {
                                                    TaskName = "SET_CLAMP_OFF";
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    string ErrorMessage = "";
                                    
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);
                                    param.Add("@Arm",Arm);

                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }

                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "MODE":
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("ALL"))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:
                                                if ((cmd.Parameter[i].Equals("AUTO") || cmd.Parameter[i].Equals("MANUAL")))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle);

                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "SIGOUT":
                                try
                                {
                                    //檢查命令格式
                                    //int no = 0;
                                    string target = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                if (cmd.Parameter[i].Equals("STOWER"))
                                                {
                                                    target = "STOWER";
                                                }
                                                else if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                   int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                   cmd.Parameter[i].Replace("P", "").Length == 3)
                                                {
                                                    target = "PORT";
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:
                                                if (target.Equals("STOWER") &&
                                                    (cmd.Parameter[i].Equals("RED") ||
                                                    cmd.Parameter[i].Equals("YELLOW") ||
                                                    cmd.Parameter[i].Equals("GREEN") ||
                                                    cmd.Parameter[i].Equals("BLUE") ||
                                                    cmd.Parameter[i].Equals("BUZZER1") ||
                                                    cmd.Parameter[i].Equals("BUZZER2")))
                                                {

                                                }
                                                else if (target.Equals("PORT") &&
                                                   (cmd.Parameter[i].Equals("LOAD") ||
                                                   cmd.Parameter[i].Equals("UNLOAD") ||
                                                   cmd.Parameter[i].Equals("OP ACCESS")))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 2:
                                                if ((cmd.Parameter[i].Equals("ON") || cmd.Parameter[i].Equals("OFF") || cmd.Parameter[i].Equals("BLINK")))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }


                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle);
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "EVENT":
                                try
                                {
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                if ((cmd.Parameter[i].Equals("ALL") ||
                                                    cmd.Parameter[i].Equals("MAPDT") ||
                                                    cmd.Parameter[i].Equals("TRANSREQ") ||
                                                    cmd.Parameter[i].Equals("SYSTEM") ||
                                                    cmd.Parameter[i].Equals("PORT") ||
                                                    cmd.Parameter[i].Equals("PRS") ||
                                                    cmd.Parameter[i].Equals("FFU")))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:


                                                if ((cmd.Parameter[i].Equals("ON") || cmd.Parameter[i].Equals("OFF")))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle);


                                    switch (cmd.Parameter[0])
                                    {
                                        case "ALL":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle);
                                            break;
                                        case "MAPDT":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle);
                                            break;
                                        case "TRANSREQ":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle);
                                            break;
                                        case "SYSTEM":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle);
                                            break;
                                        case "PORT":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle);
                                            break;
                                        case "PRS":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle);
                                            break;
                                        case "FFU":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle);
                                            break;
                                        default:
                                            //命令錯誤
                                            SendNak(WaitForHandle, "Command format error.");
                                            SendInfo(WaitForHandle);
                                            break;
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "SIZE":
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                //Parameter 1 designates transfer source, and Parameter 4 designates transfer destination.
                                                if (cmd.Parameter[i].IndexOf("ARM") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ARM", ""), out no) &&
                                                    cmd.Parameter[i].Replace("ARM", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 3)
                                                {

                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                    cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {

                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {

                                                }
                                                else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("LL", ""), out no) &&
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 3))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:

                                                //Parameter 2 designates the End-EF used for carrying out. Parameter 3 designates the End-EF used for carrying in.
                                                if ((cmd.Parameter[i].Equals("200") || cmd.Parameter[i].Equals("300")))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle);



                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            default:
                                //命令錯誤
                                SendNak(WaitForHandle, "Command format error.");
                                SendInfo(WaitForHandle);
                                break;
                        }
                        break;
                    #endregion
                    #region MOV

                    case CommandType.MOV:
                        switch (cmd.Command.ToUpper())
                        {
                            case "INIT":
                                //int no = 0;
                                //初始化
                                try
                                {
                                    //檢查命令格式

                                    string TaskName = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                TaskName = cmd.Parameter[i];
                                                //If the parameter is omitted, the system acts in the same manner as when "ALL" is designated.
                                                if (cmd.Parameter[i].Equals("ALL"))
                                                {


                                                }
                                                else if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                   int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                   cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {



                                                }
                                                else if (cmd.Parameter[i].IndexOf("ROB") != -1 &&
                                                  int.TryParse(cmd.Parameter[i].Replace("ROB", ""), out no) &&
                                                   cmd.Parameter[i].Replace("ROB", "").Length == 1)
                                                {


                                                }
                                                else if (cmd.Parameter[i].Equals("ROB"))
                                                {
                                                    TaskName = "ROB1";
                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                   cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {


                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    TaskName = "ALIGN1";
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }
                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName += "_Init";
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "ORGSH":
                                //原點復歸 
                                // int no = 0;
                                try
                                {
                                    string TaskName = "";
                                    string Target = "";
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                
                                                
                                                //If the parameter is omitted, the system acts in the same manner as when "ALL" is designated.
                                                if (cmd.Parameter[i].Equals("ALL"))
                                                {
                                                    TaskName = cmd.Parameter[i];
                                                }
                                                else if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                   int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                   cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {
                                                    TaskName = "LOADPORT";
                                                    Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");
                                                }
                                                else if (cmd.Parameter[i].IndexOf("ROB") != -1 &&
                                                  int.TryParse(cmd.Parameter[i].Replace("ROB", ""), out no) &&
                                                   cmd.Parameter[i].Replace("ROB", "").Length == 1)
                                                {
                                                    TaskName = "ROBOT";
                                                    Target = NodeNameConvert(cmd.Parameter[i], "ROBOT");
                                                }
                                                else if (cmd.Parameter[i].Equals("ROB"))
                                                {
                                                    TaskName = "ROBOT";
                                                    Target = NodeNameConvert("ROB1", "ROBOT");                                                    
                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                   cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {
                                                    TaskName = "ALIGNER";
                                                    Target = NodeNameConvert(cmd.Parameter[i], "ALIGNER");
                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    TaskName = "ALIGNER";
                                                    Target = NodeNameConvert("ALIGN1", "ALIGNER");                                                    
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }
                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName += "_ORGSH";

                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);
                                    

                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }

                                break;
                            case "LOCK":
                                //LoadPort Clamp lock
                                try
                                {

                                    string TaskName = "";
                                    string Target = "";
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                //TaskName = cmd.Parameter[i];
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i],"LOADPORT");
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "LOCK";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);

                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "UNLOCK":
                                //LoadPort Clamp unlock
                                //If FOUP is not closed, closes FOUP.
                                //If FOUP is not moved to undock position, unlocks FOUP clamp after moving FOUP to undock position.

                                try
                                {
                                    string TaskName = "";
                                    string Target = "";
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                //TaskName = cmd.Parameter[i];
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "UNLOCK";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "DOCK":
                                //Moves clamped FOUP to dock position.
                                //If FOUP is not clamped, moves FOUP to dock position after clamping it.
                                try
                                {
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Target = "";
                                
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                //TaskName = cmd.Parameter[i];
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "DOCK";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "UNDOCK":
                                //Moves FOUP to undock position.
                                //If FOUP is not closed, moves FOUP to undock position after closing it.                               
                                try
                                {
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Target = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //TaskName = cmd.Parameter[i];
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "UNDOCK";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "OPEN":

                                /*Opens FOUP door and performs wafer mapping at the same time.
                                  If FOUP clamp is not locked, execute locked FOUP clamp.
                                  If FOUP is not moved to dock position, FOUP door is opened after moving FOUP to dock position.*/
                                try
                                {
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Target = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                TaskName = cmd.Parameter[i];
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "OPEN";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }

                                break;
                            case "CLOSE":
                                //Closes FOUP door to dock position
                                try
                                {
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Target = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //TaskName = cmd.Parameter[i];
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "CLOSE";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "WAFSH":
                                //Performs wafer mapping for opened FOUP.
                                try
                                {
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Target = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //TaskName = cmd.Parameter[i];
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "WAFSH";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "GOTO":
                                //Moves the robot to the front of the designated object.
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Slot = "";
                                    string Target = "";
                                    string Position = "";
                                    string Arm = "";

                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                //Parameter 1 designates transfer source, and Parameter 4 designates transfer destination.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 3)
                                                {
                                                    Target = cmd.Parameter[i].Substring(0, 2);
                                                    Slot = no.ToString().Substring(1);
                                                    Position = NodeNameConvert(Target, "LOADPORT");
                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no))
                                                {
                                                    Target = cmd.Parameter[i];
                                                    Slot = "1";
                                                    Position = NodeNameConvert(Target, "ALIGNER");
                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    Target = "ALIGN1";
                                                    Position = NodeNameConvert(Target, "ALIGNER");
                                                }
                                                else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 1 ||
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 3) && int.TryParse(cmd.Parameter[i].Substring(3), out no)))
                                                {
                                                    Target = cmd.Parameter[i].Substring(0, 3);
                                                    Position = NodeNameConvert(Target, "STAGE");
                                                    if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    {
                                                        Slot = no.ToString();
                                                    }
                                                    else
                                                    {
                                                        Slot = "1";
                                                    }
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                //TaskName = Target;
                                                break;
                                            case 1:
                                                //TaskName += cmd.Parameter[i];
                                                //Parameter 2 designates the End-EF used for carrying out. Parameter 3 designates the End-EF used for carrying in.
                                                if ((cmd.Parameter[i].Equals("ARM1")) || cmd.Parameter[i].Equals("ARM3"))
                                                {
                                                    Arm = "1";
                                                }
                                                else if (cmd.Parameter[i].Equals("ARM2"))
                                                {
                                                    Arm = "2";
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            case 2:
                                                TaskName += cmd.Parameter[i];
                                                //Parameter 2 designates the End-EF used for carrying out. Parameter 3 designates the End-EF used for carrying in.
                                                if ((cmd.Parameter[i].Equals("UP") || cmd.Parameter[i].Equals("DOWN")))
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName += "_GOTO";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Slot", Slot);
                                    param.Add("@Arm", Arm);
                                    param.Add("@Position", Position);
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }

                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "LOAD":
                                //The robot carries out a wafer from designated position.
                                try
                                {
                                    // int no = 0;
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Slot = "";
                                    string Target = "";
                                    string Position = "";
                                    string Arm = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:

                                                //Parameter 1 designates transfer source, and Parameter 4 designates transfer destination.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 3)
                                                {
                                                    Target = cmd.Parameter[i].Substring(0, 2);
                                                    Slot = no.ToString().Substring(1);
                                                    Position = NodeNameConvert(Target, "LOADPORT");
                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no))
                                                {
                                                    Target = cmd.Parameter[i];
                                                    Slot = "1";
                                                    Position = NodeNameConvert(Target, "ALIGNER");
                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    Target = "ALIGN1";
                                                    Slot = "1";
                                                    Position = NodeNameConvert(Target, "ALIGNER");
                                                }
                                                else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 1 ||
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 3 && int.TryParse(cmd.Parameter[i].Substring(3), out no))))
                                                {
                                                    Target = cmd.Parameter[i].Substring(0, 3);
                                                    Position = NodeNameConvert(Target, "STAGE");
                                                    if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    {
                                                        Slot = no.ToString();
                                                    }
                                                    else
                                                    {
                                                        Slot = "1";
                                                    }
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                //TaskName = Target;
                                                break;
                                            case 1:
                                                //TaskName += cmd.Parameter[i];
                                                //Parameter 2 designates the End-EF used for carrying out. Parameter 3 designates the End-EF used for carrying in.
                                                if ((cmd.Parameter[i].Equals("ARM1")))
                                                {
                                                    Arm = "1";
                                                }
                                                else if (cmd.Parameter[i].Equals("ARM2"))
                                                {
                                                    Arm = "2";
                                                }
                                                else if (cmd.Parameter[i].Equals("ARM3"))
                                                {
                                                    Arm = "3";
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "LOAD";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Slot", Slot);
                                    param.Add("@Arm", Arm);
                                    param.Add("@Position", Position);
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "UNLOAD":
                                //The robot carries a wafer in designated position.
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Slot = "";
                                    string Target = "";
                                    string Position = "";
                                    string Arm = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //TaskName = cmd.Parameter[i];
                                                //Parameter 1 designates transfer source, and Parameter 4 designates transfer destination.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 3)
                                                {
                                                    Target = cmd.Parameter[i].Substring(0, 2);
                                                    Slot = no.ToString().Substring(1);
                                                    Position = NodeNameConvert(Target, "LOADPORT");
                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no))
                                                {
                                                    Target = cmd.Parameter[i];
                                                    Slot = "1";
                                                    Position = NodeNameConvert(Target, "ALIGNER");
                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    Target = "ALIGN1";
                                                    Slot = "1";
                                                    Position = NodeNameConvert(Target, "ALIGNER");
                                                }
                                                else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 1 ||
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 3 && int.TryParse(cmd.Parameter[i].Substring(3), out no))))
                                                {
                                                    Target = cmd.Parameter[i].Substring(0, 3);
                                                    Position = NodeNameConvert(Target, "STAGE");
                                                    if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    {
                                                        Slot = no.ToString();
                                                    }
                                                    else
                                                    {
                                                        Slot = "1";
                                                    }
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:
                                                //TaskName += cmd.Parameter[i];
                                                //Parameter 2 designates the End-EF used for carrying out. Parameter 3 designates the End-EF used for carrying in.
                                                if ((cmd.Parameter[i].Equals("ARM1")))
                                                {
                                                    Arm = "1";
                                                }
                                                else if (cmd.Parameter[i].Equals("ARM2"))
                                                {
                                                    Arm = "2";
                                                }
                                                else if (cmd.Parameter[i].Equals("ARM3"))
                                                {
                                                    Arm = "3";
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "UNLOAD";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Slot", Slot);
                                    param.Add("@Arm", Arm);
                                    param.Add("@Position", Position);
                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "TRANS":
                                //Designates the transfer source and transfer destination to transfer a wafer.
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    string TaskName = "";
                                    string FromSlot = "";
                                    string FromARM = "";
                                    string FromTarget = "";
                                    string ToSlot = "";
                                    string ToARM = "";
                                    string ToTarget = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                            case 3:

                                                //Parameter 1 designates transfer source, and Parameter 4 designates transfer destination.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 3)
                                                {
                                                    if (i == 0)
                                                    {
                                                        FromTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 2), "LOADPORT");
                                                        FromSlot = no.ToString().Substring(1);
                                                    }
                                                    else if (i == 3)
                                                    {
                                                        ToTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 2), "LOADPORT");
                                                        ToSlot = no.ToString().Substring(1);
                                                    }
                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no))
                                                {
                                                    if (i == 0)
                                                    {
                                                        FromTarget = NodeNameConvert(cmd.Parameter[i], "ALIGNER");
                                                        FromSlot = "1";
                                                    }
                                                    else if (i == 3)
                                                    {
                                                        ToTarget = NodeNameConvert(cmd.Parameter[i], "ALIGNER");
                                                        ToSlot = "1";
                                                    }
                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    if (i == 0)
                                                    {
                                                        FromTarget = NodeNameConvert("ALIGN1", "ALIGNER");
                                                        FromSlot = "1";
                                                    }
                                                    else if (i == 3)
                                                    {
                                                        ToTarget = NodeNameConvert("ALIGN1", "ALIGNER");
                                                        ToSlot = "1";
                                                    }
                                                }
                                                else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 1 ||
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 3 && int.TryParse(cmd.Parameter[i].Substring(3), out no))))
                                                {
                                                    if (i == 0)
                                                    {
                                                        if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                        {
                                                            FromTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                            FromSlot = no.ToString();
                                                        }
                                                        else
                                                        {
                                                            FromTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                            FromSlot = "1";
                                                        }

                                                    }
                                                    else if (i == 3)
                                                    {
                                                        if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                        {
                                                            ToTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                            ToSlot = no.ToString();
                                                        }
                                                        else
                                                        {
                                                            ToTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                            ToSlot = "1";
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:
                                            case 2:

                                                //Parameter 2 designates the End-EF used for carrying out. Parameter 3 designates the End-EF used for carrying in.
                                                if ((cmd.Parameter[i].Equals("ARM1") || cmd.Parameter[i].Equals("ARM2") || cmd.Parameter[i].Equals("ARM3")))
                                                {
                                                    if (i == 1)
                                                    {
                                                        FromARM = cmd.Parameter[i];
                                                    }
                                                    else if (i == 2)
                                                    {
                                                        ToARM = cmd.Parameter[i];
                                                    }
                                                }
                                                if ((cmd.Parameter[i].Equals("ARM1")))
                                                {
                                                    if (i == 1)
                                                    {
                                                        FromARM = "1";
                                                    }
                                                    else if (i == 2)
                                                    {
                                                        ToARM = "1";
                                                    }
                                                }
                                                else if (cmd.Parameter[i].Equals("ARM2"))
                                                {
                                                    if (i == 1)
                                                    {
                                                        FromARM = "2";
                                                    }
                                                    else if (i == 2)
                                                    {
                                                        ToARM = "2";
                                                    }
                                                }
                                                else if (cmd.Parameter[i].Equals("ARM3"))
                                                {
                                                    if (i == 1)
                                                    {
                                                        FromARM = "3";
                                                    }
                                                    else if (i == 2)
                                                    {
                                                        ToARM = "3";
                                                    }
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }

                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查  FromTarget + FromARM + ToTarget + ToARM +

                                    string ErrorMessage = "";
                                    TaskName = "TRANS";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@FromTarget", FromTarget);
                                    param.Add("@ToTarget", ToTarget);
                                    param.Add("@FromARM", FromARM);
                                    param.Add("@ToARM", ToARM);
                                    param.Add("@FromSlot", FromSlot);
                                    param.Add("@ToSlot", ToSlot);

                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }

                                break;
                            case "CHANGE":
                                //A wafer is exchanged at designated position.
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Target = "";
                                    string TargetSlot = "";
                                    string CarryOutARM = "";
                                    string CarryInARM = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates a destination to exchange a wafer
                                                if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                    cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "ALIGNER");
                                                    TargetSlot = "1";
                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    Target = NodeNameConvert("ALIGN1", "ALIGNER");
                                                    TargetSlot = "1";
                                                }
                                                else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 1 ||
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 3 && int.TryParse(cmd.Parameter[i].Substring(3), out no))))
                                                {
                                                    if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    {
                                                        Target = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                        TargetSlot = no.ToString();
                                                    }
                                                    else
                                                    {
                                                        Target = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                        TargetSlot = "1";
                                                    }
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            case 1:
                                            case 2:
                                                //Parameter 2 designates the End-EF used for carrying out. Parameter 3 designates the End-EF used for carrying in.
                                                if ((cmd.Parameter[i].Equals("ARM1") || cmd.Parameter[i].Equals("ARM2")))
                                                {
                                                    if (i == 1)
                                                    {
                                                        CarryOutARM = cmd.Parameter[i].Replace("ARM", "");
                                                    }
                                                    else if (i == 2)
                                                    {
                                                        CarryInARM = cmd.Parameter[i].Replace("ARM", "");
                                                    }
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "CHANGE";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", Target);
                                    param.Add("@Slot", TargetSlot);
                                    param.Add("@CarryOutARM", CarryOutARM);
                                    param.Add("@CarryInARM", CarryInARM);

                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }

                                break;
                            case "ALIGN":
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Slot = "";
                                    string Target = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                    cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {
                                                    Target = cmd.Parameter[i];
                                                    Slot = "1";
                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    Target = "ALIGN1";
                                                    Slot = "1";
                                                }
                                                else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("LL", ""), out no) &&
                                                    (cmd.Parameter[i].Replace("LL", "").Length == 1 || cmd.Parameter[i].Replace("LL", "").Length == 3))
                                                {
                                                    if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    {
                                                        Target = cmd.Parameter[i].Substring(0, 2);
                                                        Slot = no.ToString().Substring(1);
                                                    }
                                                    else
                                                    {
                                                        Target = cmd.Parameter[i].Substring(0, 2);
                                                        Slot = "1";
                                                    }
                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查

                                    string ErrorMessage = "";
                                    TaskName = "ALIGN";
                                    Dictionary<string, string> param = new Dictionary<string, string>();
                                    param.Add("@Target", NodeNameConvert(Target, "ALIGNER"));

                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }

                                break;
                            case "HOME":
                                //Moves each unit to the home position. (waiting position before transfer)
                                try
                                {
                                    // int no = 0;
                                    //檢查命令格式
                                    string TaskName = "";
                                    string Target = "";
                                    string Method = "";
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("ROB") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("ROB", ""), out no) &&
                                                    cmd.Parameter[i].Replace("ROB", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "ROBOT");
                                                    Method = "RobotHome";
                                                }
                                                else if (cmd.Parameter[i].Equals("ROB"))
                                                {
                                                    Target = NodeNameConvert("ROB1", "ROBOT");
                                                    Method = "RobotHome";
                                                }
                                                else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                   int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                   cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "ALIGNER");
                                                    Method = "AlignerHome";
                                                }
                                                else if (cmd.Parameter[i].Equals("ALIGN"))
                                                {
                                                    Target = NodeNameConvert("ALIGN1", "ALIGNER");
                                                    Method = "AlignerHome";
                                                }
                                                else if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {
                                                    Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    string ErrorMessage = "";
                                    if (Target.IndexOf("LOADPORT") != -1)
                                    {
                                        TaskName = "LOADPORT_Init";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);
                                    }
                                    else
                                    {
                                        TaskName = "HOME";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        param.Add("@Method", Method);

                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);
                                    }
                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "HOLD":
                                //Stops motion axes of EFEM temporarily. (Stop upon deceleration).
                                //This is invalid for the load port operation.
                                try
                                {
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }
                                    }
                                    //通過檢查
                                   
                                    string ErrorMessage = "";
                                    string TaskName = "HOLD";
                                    //Dictionary<string, string> param = new Dictionary<string, string>();


                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "RESTR":
                                //Resumes motion stopped temporarily by "Hold" message.
                                try
                                {
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }
                                    }
                                    //通過檢查
                                  
                                    string ErrorMessage = "";
                                    string TaskName = "RESTR";
                                    //Dictionary<string, string> param = new Dictionary<string, string>();


                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "ABORT":
                                //Aborts and ends the motion stopped temporarily by "HOLD" message.
                                //If the End - EF is being extended, the motion is stopped after returning the End - EF to origin.
                                try
                                {
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }
                                    }
                                    //通過檢查
                                   
                                    string ErrorMessage = "";
                                    string TaskName = "ABORT";
                                    //Dictionary<string, string> param = new Dictionary<string, string>();


                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "EMS":
                                //Performs emergency stop of EFEM.
                                //Continuing or retrying operation is not available. If emergency stop of EFEM is performed by this message, EFEM must
                                //be initialized from "(5) Start EFEM initialization" in "5-1 SETUP sequence".
                                try
                                {
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }
                                    }
                                    //通過檢查
                                   
                                    string ErrorMessage = "";
                                    string TaskName = "EMS";
                                    //Dictionary<string, string> param = new Dictionary<string, string>();


                                    TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                    if (!ErrorMessage.Equals(""))
                                    {
                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        SendAck(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "TRANSREQ":
                                //Requests for E84 automatic transfer, or acquires current request status for each port.
                                //int no = 0;
                                //檢查命令格式
                                for (int i = 0; i < cmd.Parameter.Count; i++)
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            //Designates a destination to exchange a wafer
                                            if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                cmd.Parameter[i].Replace("P", "").Length == 1)
                                            {

                                            }
                                            else
                                            {
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                            }
                                            break;
                                        case 1:
                                            //Parameter 2 designates the End-EF used for carrying out. Parameter 3 designates the End-EF used for carrying in.
                                            if (cmd.Parameter[i].Equals("LOAD") || cmd.Parameter[i].Equals("UNLOAD") || cmd.Parameter[i].Equals("STOP"))
                                            {

                                            }
                                            else
                                            {
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                            }
                                            break;
                                        default:
                                            //命令錯誤
                                            SendNak(WaitForHandle, "Command format error.");
                                            SendInfo(WaitForHandle);
                                            return;
                                    }

                                }
                                //通過檢查
                                SendAck(WaitForHandle);
                                SendInfo(WaitForHandle);
                                break;
                            case "ADPLOCK":
                                //Locks OC adapter on a load port.
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle);
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "ADPUNLOCK":
                                //Unlocks OC adapter on a load port.
                                try
                                {
                                    //int no = 0;
                                    //檢查命令格式
                                    for (int i = 0; i < cmd.Parameter.Count; i++)
                                    {
                                        switch (i)
                                        {
                                            case 0:
                                                //Designates aligner.
                                                if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                    int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                    cmd.Parameter[i].Replace("P", "").Length == 1)
                                                {

                                                }
                                                else
                                                {
                                                    //命令錯誤
                                                    SendNak(WaitForHandle, "Command format error.");
                                                    SendInfo(WaitForHandle);
                                                    return;
                                                }
                                                break;
                                            default:
                                                //命令錯誤
                                                SendNak(WaitForHandle, "Command format error.");
                                                SendInfo(WaitForHandle);
                                                return;
                                        }

                                    }
                                    //通過檢查
                                    SendAck(WaitForHandle);
                                    SendInfo(WaitForHandle);
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            default:
                                //命令錯誤
                                SendNak(WaitForHandle, "Command format error.");
                                SendInfo(WaitForHandle);
                                break;
                        }
                        break;
                        #endregion
                }



            }
            catch (Exception e)
            {
                _EventReport.On_CommandMessage(e.StackTrace);
            }
        }

        private string NodeNameConvert(string Param, string NodeType)
        {
            string result = "";
            Param = Param.ToUpper();
            switch (NodeType.ToUpper())
            {
                case "LOADPORT":
                    int PortNo = Convert.ToInt16(Param.Replace("P", "").Substring(0, 1));
                    result = "LOADPORT" + PortNo.ToString("00");
                    break;
                case "ALIGNER":
                    int AlignerNo = 0;
                    if (Param.Equals("ALIGN"))
                    {
                        AlignerNo = 1;
                    }
                    else
                    {
                        AlignerNo = Convert.ToInt16(Param.Replace("ALIGN", ""));
                    }
                    result = "ALIGNER" + AlignerNo.ToString("00");
                    break;
                case "ROBOT":
                    int RobotNo = 0;
                    if (Param.Equals("ROB"))
                    {
                        RobotNo = 1;
                    }
                    else
                    {
                        RobotNo = Convert.ToInt16(Param.Replace("ROB", ""));
                    }
                    result = "ROBOT" + RobotNo.ToString("00");
                    break;
                case "STAGE":
                    string Stage = Param.Substring(2, 1);
                    result = "PM" + Stage;
                    break;

            }

            return result;
        }

        public void On_Handling_TimeOut(OnHandling TimeOutCmd)
        {
            string key = TimeOutCmd.Cmd.Command;
            if (TimeOutCmd.INF_RetryCount < 3)
            {
                TimeOutCmd.INF_RetryCount++;

                Comm.Send(TimeOutCmd.Handler, TimeOutCmd.NotConfirmMsg);
                _EventReport.On_CommandMessage("Send:" + TimeOutCmd.NotConfirmMsg);
            }
            else
            {
                TimeOutCmd.SetTimeOutMonitor(false);//設定Timeout監控停止
                _EventReport.On_CommandMessage("Retry Timeout:" + TimeOutCmd.NotConfirmMsg);
                OnHandlingCmds.TryRemove(key, out TimeOutCmd);//從待處理名單移除

            }
        }

        #region Event report from Transfer control
        //***************Event report from Transfer control*****************Begin

        public void On_TaskJob_Finished(string TaskID)
        {
            OnHandling WaitForHandle;
            if (OnHandlingCmds.TryGetValue(TaskID, out WaitForHandle))
            {
                SendInfo(WaitForHandle);
            }
            else
            {
                logger.Error("On_TaskJob_Aborted 找不到 TaskID:" + TaskID);
            }
        }

        public void On_TaskJob_Aborted(string TaskID, string Message)
        {
            OnHandling WaitForHandle;
            if (OnHandlingCmds.TryGetValue(TaskID, out WaitForHandle))
            {

                SendABS(WaitForHandle, "TEST");
            }
            else
            {
                logger.Error("On_TaskJob_Aborted 找不到 TaskID:" + TaskID + " Message:" + Message);
            }
        }



        //***************Event report from Transfer control*****************End
        #endregion
    }
}
