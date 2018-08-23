using EFEMInterface.Comm;
using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using TransferControl.Engine;
using TransferControl.Management;

namespace EFEMInterface.MessageInterface
{
    public class RorzeInterface : ICommMessage, IHandlingTimeOutReport, IHostInterfaceReport, IUserInterfaceReport
    {
        ILog logger = LogManager.GetLogger(typeof(RorzeInterface));

        private Dictionary<string, OnHandling> OnHandlingCmds = new Dictionary<string, OnHandling>();

        IEFEMControl _EventReport;
        RouteControl RTCtrl;

        SocketServer Comm;

        public RorzeInterface(IEFEMControl EventReport)
        {
            _EventReport = EventReport;
            Comm = new SocketServer(this);
            RTCtrl = new RouteControl(this, this);
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
                    //result += "/" + cmd.Parameter[0] + "/" + data;
                    result += "/" + cmd.Parameter[0];
                    if (!data1.Equals(""))
                    {
                        result += "/" + data1;
                    }
                    break;
                case "TRANSREQ":
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

        private void SendABS(OnHandling WaitForHandle, string param1, string param2, string detail)
        {
            string ErrorMsg = ErrorAssembler(WaitForHandle.Cmd, param1, param2);
            Comm.Send(WaitForHandle.Handler, ErrorMsg);
            WaitForHandle.NotConfirmMsg = ErrorMsg;
            _EventReport.On_CommandMessage("Err :" + detail);
            _EventReport.On_CommandMessage("Send:" + ErrorMsg);
            WaitForHandle.SetTimeOutMonitor(true);//設定Timeout監控開始，5秒後
        }

        private void SendCancel(OnHandling WaitForHandle, string Factor, string Place, string detail)
        {
            //回報設備不可使用
            string CancelMsg = CancelAssembler(WaitForHandle.Cmd, Factor, Place);
            Comm.Send(WaitForHandle.Handler, CancelMsg);
            _EventReport.On_CommandMessage("Err :" + detail);
            _EventReport.On_CommandMessage("Send:" + CancelMsg);
        }

        private void SendInfo(OnHandling WaitForHandle)
        {
            string CommandMsg = CmdAssembler(WaitForHandle.Cmd, CommandType.INF);//傳送動作完成給上位系統
            WaitForHandle.NotConfirmMsg = CommandMsg;
            Comm.Send(WaitForHandle.Handler, CommandMsg);
            _EventReport.On_CommandMessage("Send:" + CommandMsg);
            WaitForHandle.SetTimeOutMonitor(true);//設定Timeout監控開始，5秒後
        }

        private void SendInfo(OnHandling WaitForHandle, string data1, string data2)
        {
            string CommandMsg = InfoAssembler(WaitForHandle.Cmd, data1, data2);//回傳資料給上位系統
            WaitForHandle.NotConfirmMsg = CommandMsg;
            Comm.Send(WaitForHandle.Handler, CommandMsg);
            _EventReport.On_CommandMessage("Send:" + CommandMsg);
            WaitForHandle.SetTimeOutMonitor(true);//設定Timeout監控開始，5秒後
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
            OnHandlingCmds.Add(WaitForHandle.ID, WaitForHandle);

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
                        OnHandlingCmds.Add(WaitForHandle.ID, WaitForHandle);


                        break;
                    case CommandType.ACK://收到上位系統回覆
                        List<OnHandling> tmp = OnHandlingCmds.Values.ToList();
                        tmp.Sort((x, y) => { return x.ReceiveTime.CompareTo(y.ReceiveTime); });

                        var findHandling = from Handling in tmp
                                           where Handling.Cmd.Command.Equals(cmd.Command)
                                           select Handling;

                        if (findHandling.Count() != 0)
                        {
                            WaitForHandle = findHandling.First();
                            WaitForHandle.SetTimeOutMonitor(false);//設定Timeout監控停止
                            OnHandlingCmds.Remove(WaitForHandle.ID);//從待處理名單移除
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
                    case CommandType.GET:
                        switch (cmd.Command.ToUpper())
                        {
                            case "MAPDT":
                                //取得LoadPort Mapping 結果
                                try
                                {
                                    node = NodeManagement.Get(NodeNameConvert(cmd.Parameter[0], "LOADPORT"));
                                }
                                catch
                                {
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                    break;
                                }
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

                                    On_Command_Excuted(node, txn, Msg);
                                    //*********************test   end*********************
                                }
                                else
                                {
                                    //回報設備不可使用
                                    SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "Loadport not found.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                            case "ERROR":
                                //取得ERROR狀態

                                //*********************test begin*********************
                                SendAck(WaitForHandle);
                                SendInfo(WaitForHandle, "COMMAND", "ROBOT");

                                //*********************test   end*********************

                                break;
                            case "CLAMP":
                                //取得CLAMP狀態
                                try
                                {

                                    if (cmd.Parameter[0].IndexOf("ARM") != -1)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "OFF");
                                    }
                                    else if (cmd.Parameter[0].IndexOf("ALIGN") != -1)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "ON");
                                    }
                                    else
                                    {
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
                            case "STATE":
                                //取得STATE狀態

                                //*********************test begin*********************
                                try
                                {
                                    if (cmd.Parameter[0].Equals("VER"))
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, "VER", "1.0.0.1(2018-08-01)");
                                    }
                                    else if (cmd.Parameter[0].Equals("TRACK"))
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, "TRACK", "NONE/200/300");
                                    }
                                    else if (cmd.Parameter[0].IndexOf("PRS") != -1 && cmd.Parameter[0].Replace("PRS", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "SNO1|00000000,SNO2|00003000,SNO3|00009527,SNO4|88888888");
                                    }
                                    else if (cmd.Parameter[0].IndexOf("FFU") != -1 && cmd.Parameter[0].Replace("FFU", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "FNO1|00000000,FNO2|00003000,FNO3|00009527,FNO4|88888888");
                                    }
                                    else
                                    {
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
                                //*********************test   end*********************

                                break;
                            case "MODE":
                                //取得MODE狀態

                                //*********************test begin*********************
                                try
                                {
                                    if (cmd.Parameter[0].IndexOf("P") != -1 && cmd.Parameter[0].Replace("P", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "MANUAL");
                                    }
                                    else
                                    {
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
                                //*********************test   end*********************

                                break;
                            case "TRANSREQ":
                                //取得E84狀態

                                //*********************test begin*********************
                                SendAck(WaitForHandle);
                                SendInfo(WaitForHandle, "STOP", "");
                                //*********************test   end*********************

                                break;
                            case "SIGSTAT":
                                //取得SIGSTAT狀態

                                //*********************test begin*********************
                                SendAck(WaitForHandle);
                                SendInfo(WaitForHandle, "11111111111111111111111111111111", "11111111111111111111111111111111");

                                //*********************test   end*********************

                                break;
                            case "EVENT":
                                //取得EVENT狀態

                                //*********************test begin*********************
                                try
                                {
                                    switch (cmd.Parameter[0])
                                    {
                                        case "MAPDT":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle, cmd.Parameter[0], "OFF");
                                            break;
                                        case "TRANSREQ":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle, cmd.Parameter[0], "OFF");
                                            break;
                                        case "SYSTEM":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle, cmd.Parameter[0], "OFF");
                                            break;
                                        case "PORT":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle, cmd.Parameter[0], "OFF");
                                            break;
                                        case "PRS":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle, cmd.Parameter[0], "OFF");
                                            break;
                                        case "FFU":
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle, cmd.Parameter[0], "OFF");
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
                                //*********************test   end*********************

                                break;
                            case "CSTID":
                                //取得CSTID

                                //*********************test begin*********************
                                try
                                {
                                    if (cmd.Parameter[0].IndexOf("P") != -1 && cmd.Parameter[0].Replace("P", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "FOUPIDXX");
                                    }
                                    else
                                    {
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
                                //*********************test   end*********************

                                break;
                            case "SIZE":
                                //取得Wafer Size

                                //*********************test begin*********************
                                try
                                {
                                    if (cmd.Parameter[0].IndexOf("ARM") != -1 && cmd.Parameter[0].Replace("ARM", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "????");
                                    }
                                    else if (cmd.Parameter[0].IndexOf("P") != -1 && cmd.Parameter[0].Replace("P", "").Length >= 3)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "NONE");
                                    }
                                    else if (cmd.Parameter[0].IndexOf("ALIGN") != -1 && cmd.Parameter[0].Replace("ALIGN", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "200");
                                    }
                                    else if (cmd.Parameter[0].IndexOf("LL") != -1 && cmd.Parameter[0].Replace("LL", "").Length >= 3)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, cmd.Parameter[0], "300");
                                    }
                                    else
                                    {
                                        //命令錯誤
                                        SendNak(WaitForHandle, "Command format error.");
                                        SendInfo(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                //*********************test   end*********************

                                break;
                        }
                        break;
                    case CommandType.SET:
                        switch (cmd.Command.ToUpper())
                        {
                            case "ALIGN":
                                //設定Aligner旋轉Notch角度
                                try
                                {
                                    node = NodeManagement.Get(NodeNameConvert(cmd.Parameter[0], "ALIGNER"));//取得Aligner
                                }
                                catch
                                {
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                    break;
                                }
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
                                break;
                            case "ERROR":
                                try
                                {
                                    if (cmd.Parameter[0].Equals("CLEAR"))
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        //命令錯誤
                                        SendNak(WaitForHandle, "Command format error.");
                                        SendInfo(WaitForHandle);
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

                                    if (cmd.Parameter[0].IndexOf("ARM") != -1)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else if (cmd.Parameter[0].IndexOf("ALIGN") != -1)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
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
                            case "MODE":
                                try
                                {
                                    if (cmd.Parameter[0].IndexOf("P") != -1 && cmd.Parameter[0].Replace("P", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else if(cmd.Parameter[0].Equals("ALL"))
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
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
                            case "SIGOUT":
                                try
                                {
                                    if (cmd.Parameter[0].IndexOf("P") != -1 && cmd.Parameter[0].Replace("P", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else if (cmd.Parameter[0].Equals("STOWER"))
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
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
                            case "EVENT":
                                try
                                {
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
                                    if (cmd.Parameter[0].IndexOf("ARM") != -1 && cmd.Parameter[0].Replace("ARM", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else if (cmd.Parameter[0].IndexOf("P") != -1 && cmd.Parameter[0].Replace("P", "").Length >= 3)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else if (cmd.Parameter[0].IndexOf("ALIGN") != -1 && cmd.Parameter[0].Replace("ALIGN", "").Length != 0)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else if (cmd.Parameter[0].IndexOf("LL") != -1 && cmd.Parameter[0].Replace("LL", "").Length >= 3)
                                    {
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle);
                                    }
                                    else
                                    {
                                        //命令錯誤
                                        SendNak(WaitForHandle, "Command format error.");
                                        SendInfo(WaitForHandle);
                                    }
                                }
                                catch
                                {
                                    //命令錯誤
                                    SendNak(WaitForHandle, "Command format error.");
                                    SendInfo(WaitForHandle);
                                }
                                break;
                        }
                        break;
                    case CommandType.MOV:
                        SendAck(WaitForHandle);
                        SendABS(WaitForHandle, "TEST", "TEST", "TEST");
                        break;
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
                    int PortNo = Convert.ToInt16(Param.Replace("P", ""));
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
                    int StageNo = Convert.ToInt16(Param.Replace("LL", ""));
                    result = "PM" + StageNo.ToString("00");
                    break;
                case "ARM":
                    int ARMNo = 0;
                    if (Param.Equals("ARM"))
                    {
                        ARMNo = 1;
                    }
                    else
                    {
                        ARMNo = Convert.ToInt16(Param.Replace("ARM", ""));
                    }
                    result = "ARM" + ARMNo.ToString("00");
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
                OnHandlingCmds.Remove(key);//從待處理名單移除
                _EventReport.On_CommandMessage("Retry Timeout:" + TimeOutCmd.NotConfirmMsg);
            }
        }

        #region Event report from Transfer control
        //***************Event report from Transfer control*****************Begin

        public void On_Command_Excuted(Node Node, Transaction Txn, SANWA.Utility.ReturnMessage Msg)
        {
            //處理邏輯結束
            OnHandling WaitHandle;
            if (OnHandlingCmds.TryGetValue(Txn.FormName, out WaitHandle))
            {
                switch (WaitHandle.Cmd.CommandType)
                {
                    case CommandType.GET:
                        if (WaitHandle.Cmd.Equals("SIGSTAT"))
                        {

                        }
                        else
                        {
                            SendInfo(WaitHandle, Msg.Value, "");
                        }
                        break;
                    case CommandType.SET:
                        SendInfo(WaitHandle);
                        break;
                }
            }
        }

        public void On_Command_Error(Node Node, Transaction Txn, SANWA.Utility.ReturnMessage Msg)
        {
            OnHandling WaitHandle;
            if (OnHandlingCmds.TryGetValue(Txn.FormName, out WaitHandle))
            {
                SendABS(WaitHandle, "TBD", "TBD", Msg.Value);
            }
        }

        public void On_Command_Finished(Node Node, Transaction Txn, SANWA.Utility.ReturnMessage Msg)
        {

        }

        public void On_Command_TimeOut(Node Node, Transaction Txn)
        {

        }

        public void On_Event_Trigger(Node Node, SANWA.Utility.ReturnMessage Msg)
        {

        }

        public void On_Controller_State_Changed(string Device_ID, string Status)
        {

        }

        public void On_Script_Finished(Node Node, string ScriptName, string FormName)
        {

        }

        public void On_Node_State_Changed(Node Node, string Status)
        {

        }

        public void On_Eqp_State_Changed(string OldStatus, string NewStatus)
        {

        }

        public void On_Port_Begin(string PortName, string FormName)
        {

        }

        public void On_Port_Finished(string PortName, string FormName)
        {

        }

        public void On_Task_Finished(string FormName, string LapsedTime, int LapsedWfCount, int LapsedLotCount)
        {

        }

        public void On_Job_Location_Changed(Job Job)
        {

        }

        public void On_InterLock_Report(Node Node, bool InterLock)
        {

        }

        public void On_Mode_Changed(string Mode)
        {

        }

        //***************Event report from Transfer control*****************End
        #endregion
    }
}
