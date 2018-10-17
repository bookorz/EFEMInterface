using DIOControl;
using EFEMInterface.Comm;
using log4net;
using Newtonsoft.Json;
using SANWA.Utility;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
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

        public ReportEvent Events = new ReportEvent();

        AlarmMapping AlmMapping = new AlarmMapping();

        AlarmMessage LastError = null;

        OnHandling EventHandling = null;

        public bool SaftyCheckByPass = true;

        //For TDK Loadport
        private List<OnHandling> WaitForExcute = new List<OnHandling>();

        private static DBUtil dBUtil = new DBUtil();

        string SIGSTAT_SYSTEM_Data1 = "00000000";
        string SIGSTAT_SYSTEM_Data2 = "00000000";
        string SIGSTAT_PORT_Data1 = "00000000";
        string SIGSTAT_PORT_Data2 = "00000000";
        string EFEM_State = "Not Initialize";

        public class ReportEvent
        {
            public void Load()
            {
                string Sql = @"SELECT MAPDT, TRANSREQ, SYSTEM, PORT, PRS, FFU FROM config_efem_event";

                DataTable dt = dBUtil.GetDataTable(Sql, null);
                string str_json = JsonConvert.SerializeObject(dt, Formatting.Indented);

                ReportEvent tmp = JsonConvert.DeserializeObject<List<ReportEvent>>(str_json)[0];
                this.MAPDT = tmp.MAPDT;
                this.TRANSREQ = tmp.TRANSREQ;
                this.SYSTEM = tmp.SYSTEM;
                this.PORT = tmp.PORT;
                this.PRS = tmp.PRS;
                this.FFU = tmp.FFU;

                
            }

            public void Save()
            {
                string Sql = @"UPDATE config_efem_event
	                            SET
		                            MAPDT={0},
		                            TRANSREQ={1},
		                            SYSTEM={2},
		                            PORT={3},
		                            PRS={4},
		                            FFU={5}";
                Sql = string.Format(Sql, Convert.ToByte(this.MAPDT), Convert.ToByte(this.TRANSREQ), Convert.ToByte(this.SYSTEM), Convert.ToByte(this.PORT), Convert.ToByte(this.PRS), Convert.ToByte(this.FFU));
                dBUtil.ExecuteNonQuery(Sql, null);


            }

            public bool MAPDT { get; set; }
            public bool TRANSREQ { get; set; }
            public bool SYSTEM { get; set; }
            public bool PORT { get; set; }
            public bool PRS { get; set; }
            public bool FFU { get; set; }
        }

        public RorzeInterface(IEFEMControl EventReport)
        {
            SaftyCheckByPass = SANWA.Utility.Config.SystemConfig.Get().SaftyCheckByPass;
            _EventReport = EventReport;
            Comm = new SocketServer(this);
           

        }

        public class RorzeCommand
        {
            public string OrgMsg = "";
            public string Command;
            public List<string> Parameter = new List<string>();
            public string CommandType = "";
            public string CommandParam = "";
            public string Target = "";
            public string Arm = "";

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

        public void Reset()
        {
            OnHandlingCmds.Clear();
        }

        private RorzeCommand CmdParser(string Msg)
        {
            RorzeCommand each = new RorzeCommand();

            each.OrgMsg = Msg;

            string[] content = Msg.Replace(";", "").Replace("\r", "").Split(':');
            each.CommandParam = content[1];


            content = Msg.Replace(";", "").Replace("\r", "").Split(':', '/', '>');

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
            if (cmd.Command.Equals("TRANS"))
            {
                result += "/" + cmd.Parameter[0] + ">" + cmd.Parameter[1] + "/" + cmd.Parameter[2] + ">" + cmd.Parameter[3];
            }
            else
            {
                foreach (string param in cmd.Parameter)
                {
                    result += "/" + param;
                }
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

            //switch (cmd.Command)
            //{
            //    case "MAPDT":
            //    case "TRANSREQ":
            //    case "CLAMP":
            //    case "STATE":
            //    case "MODE":
            //    case "EVENT":
            //    case "CSTID":
            //    case "SIZE":
            //        //result += "/" + cmd.Parameter[0] + "/" + data;
            //        result += "/" + cmd.Parameter[0];
            //        if (!data1.Equals(""))
            //        {
            //            result += "/" + data1;
            //        }
            //        break;

            //    case "SIGSTAT":
            //result += "/" + cmd.Parameter[0] + "/" + data;
            if (cmd.Parameter.Count != 0)
            {
                result += "/" + cmd.Parameter[0];
            }
            if (!data1.Equals(""))
            {
                result += "/" + data1;
            }
            if (!data2.Equals(""))
            {
                result += "/" + data2;
            }
            //        break;
            //    case "ERROR":
            //        if (!data1.Equals(""))
            //        {
            //            result += "/" + data1;
            //        }
            //        if (!data2.Equals(""))
            //        {
            //            result += "/" + data2;
            //        }
            //        break;
            //}
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
            System.Threading.Thread.Sleep(300);
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
            OnHandlingCmds.TryRemove(WaitForHandle.ID,out WaitForHandle);
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
                string d1 = data1;
                string d2 = data2;
                string CommandMsg = InfoAssembler(WaitForHandle.Cmd, d1, d2);//回傳資料給上位系統
                WaitForHandle.NotConfirmMsg = CommandMsg;
                WaitForHandle.SetTimeOutMonitor(true);//設定Timeout監控開始，5秒後
                Comm.Send(WaitForHandle.Handler, CommandMsg);
                _EventReport.On_CommandMessage("Send:" + CommandMsg);

            }
        }
        string lastEvt = "";
        private void SendEvent(OnHandling WaitForHandle, string Type, string From, string data1, string data2)
        {
            if (WaitForHandle != null)
            {
                if (lastEvt.Equals(Type + From + data1 + data2))
                {
                    return;
                }
                lastEvt = Type + From + data1 + data2;
                RorzeCommand R = new RorzeCommand();
                R.Command = Type;
                R.Parameter.Add(From);
                string CommandMsg = InfoAssembler(R, data1, data2, true);//回傳資料給上位系統
                                                                         //WaitForHandle.NotConfirmMsg = CommandMsg;
                                                                         //WaitForHandle.SetTimeOutMonitor(true);//設定Timeout監控開始，5秒後
                Comm.Send(WaitForHandle.Handler, CommandMsg);
                _EventReport.On_CommandMessage("Send:" + CommandMsg);
            }
        }

        private void SendABS(OnHandling WaitForHandle, string param1, string param2)
        {
            if (!WaitForHandle.IsReturn)
            {
                WaitForHandle.IsReturn = true;
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
            Events.Load();
            RorzeCommand CommunityActive = new RorzeCommand();
            CommunityActive.CommandType = CommandType.INF;
            CommunityActive.Command = "READY";
            CommunityActive.CommandParam = "READY/COMM";
            CommunityActive.Parameter.Add("COMM");

            OnHandling WaitForHandle = new OnHandling(this);
            WaitForHandle.Cmd = CommunityActive;
            WaitForHandle.Handler = handler;
            OnHandlingCmds.TryAdd(WaitForHandle.ID, WaitForHandle);

            SendInfo(WaitForHandle);

            _EventReport.On_Connection_Connected();
            EventHandling = WaitForHandle;
        }

        public void On_Connection_Disconnected()
        {
            _EventReport.On_Connection_Disconnected();
        }

        public void On_Connection_Error(string Msg)
        {
            _EventReport.On_CommandMessage(Msg.ToString());
        }

        static object lockObj = new object();

        public void On_Connection_Message(Socket handler, string contents)
        {
            try
            {
                string[] cmds = contents.Split(';');
                foreach (string content in cmds)
                {
                    if (content.Trim().Equals(""))
                    {
                        continue;
                    }
                    logger.Debug("EFEM Host Recieve : " + content);
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

                            if (LastError != null && cmd.CommandType.Equals(CommandType.MOV))
                            {
                                SendCancel(WaitForHandle, "ABS", LastError.Position, "Has alarm.");
                                //SendInfo(WaitForHandle);
                                return;
                            }

                            //if (EFEM_State.Equals("Not Initialize") && cmd.CommandType.Equals(CommandType.MOV))
                            //{
                            //    if (!cmd.Command.Equals("INIT"))
                            //    {
                            //        SendCancel(WaitForHandle, "NOINITCMPL", "", "Not Initialize");
                            //       // SendInfo(WaitForHandle);
                            //        return;
                            //    }
                            //}
                            //Node rob = NodeManagement.Get("ROBOT01");
                            //if (rob.State.Equals("Not Origin") && cmd.CommandType.Equals(CommandType.MOV))
                            //{
                            //    if (!cmd.Command.Equals("ORGSH")&& !cmd.Command.Equals("INIT"))
                            //    {
                            //        SendCancel(WaitForHandle, "NOORGCMPL", "ROBOT", "Not Orgin");
                            //       // SendInfo(WaitForHandle);
                            //        return;
                            //    }
                            //}

                            //Node port;

                            //if (cmd.Command.Equals("SIGOUT"))
                            //{
                            //    string[] cmdAry = cmd.OrgMsg.Replace(";\r", "").Split(new char[] { ':', ',', '/' });
                            //    if (!cmdAry[2].Equals("STOWER"))
                            //    {
                            //        lock (WaitForExcute)
                            //        {
                            //            port = NodeManagement.Get(NodeNameConvert(cmdAry[2], "LOADPORT"));
                            //            WaitForHandle.Msg = content;
                            //            WaitForHandle.Cmd.Target = NodeNameConvert(cmdAry[2], "LOADPORT");
                            //            WaitForExcute.Add(WaitForHandle);


                            //            var find = from Handling in OnHandlingCmds.Values.ToList()
                            //                       where Handling.Cmd.Target.Equals(WaitForHandle.Cmd.Target)
                            //                       select Handling;
                            //            if (find.Count() == 0)
                            //            {
                            //                WaitForHandle = WaitForExcute.First();
                            //                WaitForExcute.Remove(WaitForHandle);
                            //            }
                            //            else
                            //            {
                            //                return;
                            //            }
                            //        }
                            //    }
                            //}

                            //var findx = from Handling in OnHandlingCmds.Values.ToList()
                            //            where Handling.Cmd.Command.Equals(cmd.Command) && Handling.Cmd.CommandType.Equals(CommandType.MOV) && !Handling.Cmd.Command.Equals("OPEN") && !Handling.Cmd.Command.Equals("CLOSE") && 
                            //            !Handling.Cmd.Command.Equals("CLAMP") && !Handling.Cmd.Command.Equals("UNCLAMP") && !Handling.Cmd.Command.Equals("LOCK") && !Handling.Cmd.Command.Equals("UNLOCK") && !Handling.Cmd.Command.Equals("WAFSH")&&
                            //            !Handling.Cmd.Command.Equals("DOCK") && !Handling.Cmd.Command.Equals("UNDOCK")
                            //            select Handling;
                            OnHandlingCmds.TryAdd(WaitForHandle.ID, WaitForHandle);
                            //if (findx.Count() != 0)
                            //{

                            //    SendCancel(WaitForHandle, ErrorCategory.CancelFactor.BUSY, "DUPLICATE", "Command already exsit.");
                            //    SendInfo(WaitForHandle);
                            //    return;
                            //}


                            break;
                        case CommandType.ACK://收到上位系統回覆
                            List<OnHandling> tmp = OnHandlingCmds.Values.ToList();
                            tmp.Sort((x, y) => { return x.ReceiveTime.CompareTo(y.ReceiveTime); });

                            var findHandling = from Handling in tmp
                                               where Handling.Cmd.Command.Equals(cmd.Command) && cmd.CommandParam.IndexOf(Handling.Cmd.CommandParam) != -1
                                               select Handling;



                            if (findHandling.Count() != 0)
                            {
                                WaitForHandle = findHandling.First();
                                WaitForHandle.SetTimeOutMonitor(false);//設定Timeout監控停止
                                OnHandlingCmds.TryRemove(WaitForHandle.ID, out WaitForHandle);//從待處理名單移除
                            }

                            //foreach(OnHandling Handling in tmp)
                            //{
                            //    if (Handling.Cmd.Command.Equals(cmd.Command) && Handling.Cmd.CommandParam.Equals(cmd.CommandParam))
                            //    {
                            //        WaitForHandle = Handling;
                            //        WaitForHandle.SetTimeOutMonitor(false);//設定Timeout監控停止
                            //        OnHandlingCmds.TryRemove(WaitForHandle.ID, out WaitForHandle);//從待處理名單移除
                            //        break;
                            //    }
                            //}


                            //if (cmd.Command.Equals("SIGOUT"))
                            //{
                            //    string[] cmdAry = cmd.OrgMsg.Replace(";\r", "").Split(new char[] { ':', ',', '/' });
                            //    if (!cmdAry[2].Equals("STOWER"))
                            //    {
                            //        OnHandling WaitExcute = null;//Queue for TDK Loadport
                            //        IEnumerable<OnHandling> findWaitExcute;
                            //        lock (WaitForExcute)
                            //        {
                            //            findWaitExcute = from each in WaitForExcute
                            //                             where each.Cmd.Target.Equals(WaitForHandle.Cmd.Target)
                            //                             select each;

                            //            if (findWaitExcute.Count() != 0)
                            //            {
                            //                WaitExcute = findWaitExcute.First();
                            //                WaitForExcute.Remove(WaitExcute);


                            //            }
                            //        }
                            //        if (WaitExcute != null)
                            //        {
                            //            port = NodeManagement.Get(WaitExcute.Cmd.Target);
                            //            if (!port.IsExcuting)
                            //            {
                            //                this.On_Connection_Message(WaitExcute.Handler, WaitExcute.Msg);
                            //            }
                            //        }
                            //    }
                            //}
                            break;
                        default:
                            //命令錯誤
                            WaitForHandle = new OnHandling(this);
                            WaitForHandle.Cmd = cmd;
                            WaitForHandle.Handler = handler;

                            findHandling = from Handling in OnHandlingCmds.Values.ToList()
                                           where Handling.Cmd.Command.Equals(cmd.Command) && Handling.Cmd.CommandType.Equals(CommandType.MOV)

                                           select Handling;
                            OnHandlingCmds.TryAdd(WaitForHandle.ID, WaitForHandle);
                            if (findHandling.Count() != 0)
                            {

                                SendCancel(WaitForHandle, ErrorCategory.CancelFactor.BUSY, "DUPLICATE", "Command already exsit.");
                                //SendInfo(WaitForHandle);

                            }
                            else
                            {
                                SendNak(WaitForHandle, "Command format error.");
                                SendInfo(WaitForHandle);
                            }



                            return;
                    }

                    //處理邏輯開始
                    Node node = null;
                    string TaskName = "";
                    string Target = "";
                    string ErrorMessage = "";
                    switch (cmd.CommandType)
                    {
                        #region GET

                        case CommandType.GET:
                            switch (cmd.Command.ToUpper())
                            {
                                case "STATUS":
                                    //int no = 0;
                                    //初始化
                                    try
                                    {
                                        //檢查命令格式

                                        TaskName = "";
                                        Target = "";
                                        for (int i = 0; i < cmd.Parameter.Count; i++)
                                        {
                                            switch (i)
                                            {
                                                case 0:
                                                    TaskName = cmd.Parameter[i];
                                                    //If the parameter is omitted, the system acts in the same manner as when "ALL" is designated.
                                                    if (cmd.Parameter[i].Equals("ALL"))
                                                    {
                                                        TaskName = "ALL";

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
                                        string Result = "";
                                        if (TaskName.Equals("ALL"))
                                        {
                                            Result = EFEM_State;
                                        }
                                        else
                                        {
                                            Node t = NodeManagement.Get(Target);
                                            Result = t.State;
                                        }

                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle,Result,"");
                                       
                                        
                                    }
                                    catch
                                    {
                                        //命令錯誤
                                        SendNak(WaitForHandle, "Command format error.");
                                        SendInfo(WaitForHandle);
                                    }
                                    break;
                                case "MAPDT":
                                    //取得LoadPort Mapping 結果
                                    #region MAPDT
                                    Target = "";
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
                                                        Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");
                                                        cmd.Target = Target;
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
                                        ErrorMessage = "";
                                        TaskName = "GET_MAPDT";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);

                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            SendAck(WaitForHandle);
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
                                        if (LastError != null)
                                        {
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle, LastError.Code_Group, LastError.Position);
                                        }
                                        else
                                        {
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle, "NOTHING", "UNDEFINITION");
                                        }

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
                                                        cmd.Target = "ROBOT01";
                                                        cmd.Arm = no.ToString();
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                       int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                       cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                    {
                                                        cmd.Target = NodeNameConvert(cmd.Parameter[i], "ALIGNER");
                                                        cmd.Arm = "1";
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

                                        ErrorMessage = "";
                                        TaskName = "GET_CLAMP";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", cmd.Target);

                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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
                                                        Node robot = NodeManagement.Get("ROBOT01");

                                                        returnValue = robot.R_Hold_Status + robot.L_Hold_Status;
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("PRS") != -1 &&
                                                        int.TryParse(cmd.Parameter[i].Replace("PRS", ""), out no) &&
                                                        cmd.Parameter[i].Replace("PRS", "").Length == 1)
                                                    {
                                                        //returnValue = "SNO1|00000000,SNO2|00003000,SNO3|00009527,SNO4|88888888";
                                                        returnValue = "N/A";
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("FFU") != -1 &&
                                                        int.TryParse(cmd.Parameter[i].Replace("FFU", ""), out no) &&
                                                        cmd.Parameter[i].Replace("FFU", "").Length == 1)
                                                    {
                                                        //returnValue = "FNO1|00000000,FNO2|00003000,FNO3|00009527,FNO4|88888888";
                                                        returnValue = "N/A";
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
                                        Target = "";
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
                                                        Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT");

                                                    }
                                                    else if (cmd.Parameter[i].Equals("SYSTEM"))
                                                    {
                                                        Target = "SYSTEM";
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

                                        if (Target.Equals("SYSTEM"))
                                        {
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle, SIGSTAT_SYSTEM_Data1, SIGSTAT_SYSTEM_Data2);
                                        }
                                        else
                                        {
                                            Node port = NodeManagement.Get(Target);
                                            if (port == null)
                                            {
                                                SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "Node not found.");
                                               // SendInfo(WaitForHandle);
                                            }
                                            else
                                            {

                                                SendAck(WaitForHandle);
                                                SendInfo(WaitForHandle, SIGSTAT_PORT_Data1, SIGSTAT_PORT_Data2);
                                            }
                                        }

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
                                        string result = "OFF";
                                        for (int i = 0; i < cmd.Parameter.Count; i++)
                                        {
                                            switch (i)
                                            {
                                                case 0:
                                                    //Designates aligner.
                                                    if (cmd.Parameter[i].Equals("MAPDT"))
                                                    {
                                                        if (Events.MAPDT)
                                                        {
                                                            result = "ON";
                                                        }
                                                    }
                                                    else if (cmd.Parameter[i].Equals("TRANSREQ"))
                                                    {
                                                        if (Events.TRANSREQ)
                                                        {
                                                            result = "ON";
                                                        }
                                                    }
                                                    else if (cmd.Parameter[i].Equals("SYSTEM"))
                                                    {
                                                        if (Events.SYSTEM)
                                                        {
                                                            result = "ON";
                                                        }
                                                    }
                                                    else if (cmd.Parameter[i].Equals("PORT"))
                                                    {
                                                        if (Events.PORT)
                                                        {
                                                            result = "ON";
                                                        }
                                                    }
                                                    else if (cmd.Parameter[i].Equals("PRS"))
                                                    {
                                                        if (Events.PRS)
                                                        {
                                                            result = "ON";
                                                        }
                                                    }
                                                    else if (cmd.Parameter[i].Equals("FFU"))
                                                    {
                                                        if (Events.FFU)
                                                        {
                                                            result = "ON";
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
                                        SendAck(WaitForHandle);
                                        SendInfo(WaitForHandle, result, "");

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
                                                        TaskName = "READ_LCD";
                                                        Target = NodeNameConvert(cmd.Parameter[i], "LOADPORT").Replace("LOADPORT","SMARTTAG");
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
                                        //SendAck(WaitForHandle);
                                        //SendInfo(WaitForHandle, "FOUPIDXX", "");
                                        Dictionary<string, string> Param = new Dictionary<string, string>();
                                        Param.Add("@Target", Target);                                      
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, Param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                            // SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            SendAck(WaitForHandle);
                                            LastError = null;
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
                                        SendInfo(WaitForHandle, "200", "");

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
                                case "SPEED":
                                    //檢查命令格式
                                    try
                                    {
                                        TaskName = "SPEED";
                                        Target = "";
                                        ErrorMessage = "";
                                        for (int i = 0; i < cmd.Parameter.Count; i++)
                                        {
                                            switch (i)
                                            {
                                                case 0:
                                                    //TaskName = cmd.Parameter[i];
                                                    //If the parameter is omitted, the system acts in the same manner as when "ALL" is designated.

                                                    if (cmd.Parameter[i].IndexOf("ROB") != -1 &&
                                                     int.TryParse(cmd.Parameter[i].Replace("ROB", ""), out no) &&
                                                      cmd.Parameter[i].Replace("ROB", "").Length == 1)
                                                    {
                                                        //TaskName = "ROBOT";
                                                        Target = NodeNameConvert(cmd.Parameter[i], "ROBOT");

                                                    }
                                                    else if (cmd.Parameter[i].Equals("ROB"))
                                                    {
                                                        // TaskName = "ROBOT";
                                                        Target = NodeNameConvert("ROB1", "ROBOT");
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                        int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no) &&
                                                       cmd.Parameter[i].Replace("ALIGN", "").Length == 1)
                                                    {
                                                        //TaskName = "ALIGNER";
                                                        Target = NodeNameConvert(cmd.Parameter[i], "ALIGNER");

                                                    }
                                                    else if (cmd.Parameter[i].Equals("ALIGN"))
                                                    {
                                                        //TaskName = "ALIGNER";
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
                                                    if (!int.TryParse(cmd.Parameter[i], out no))
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
                                        if (no == 100)
                                        {
                                            no = 0;
                                        }
                                        Dictionary<string, string> Param = new Dictionary<string, string>();
                                        Param.Add("@Target", Target);
                                        Param.Add("@Value", no.ToString());
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, Param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            SendAck(WaitForHandle);
                                            LastError = null;
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
                                                          //  SendInfo(WaitForHandle);
                                                        }

                                                    }
                                                    else
                                                    {
                                                        SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "ptAligner not found.");
                                                      //  SendInfo(WaitForHandle);
                                                    }
                                                }
                                                else
                                                {
                                                    SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "NextRobot not found.");
                                                   // SendInfo(WaitForHandle);
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
                                               // SendInfo(WaitForHandle);
                                            }
                                        }
                                        else
                                        {
                                            //回報設備不可使用

                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", "Aligner not found.");
                                          //  SendInfo(WaitForHandle);
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

                                        ErrorMessage = "";
                                        TaskName = "SET_ERROR";
                                        //Dictionary<string, string> param = new Dictionary<string, string>();


                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                          //  SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            SendAck(WaitForHandle);
                                            LastError = null;
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
                                        TaskName = "";
                                        Target = "";
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
                                                    if (cmd.Parameter[i].Equals("ON"))
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
                                        ErrorMessage = "";

                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        param.Add("@Arm", Arm);

                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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
                                        TaskName = "";
                                        Target = "";
                                        string type = "";
                                        string point = "";
                                        string state = "";
                                        for (int i = 0; i < cmd.Parameter.Count; i++)
                                        {
                                            switch (i)
                                            {
                                                case 0:
                                                    if (cmd.Parameter[i].Equals("STOWER"))
                                                    {
                                                        type = "STOWER";
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                       int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                       cmd.Parameter[i].Replace("P", "").Length == 1)
                                                    {
                                                        type = "PORT";
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
                                                case 1:
                                                    if (type.Equals("STOWER") &&
                                                        (cmd.Parameter[i].Equals("RED") ||
                                                        cmd.Parameter[i].Equals("YELLOW") ||
                                                        cmd.Parameter[i].Equals("GREEN") ||
                                                        cmd.Parameter[i].Equals("BLUE") ||
                                                        cmd.Parameter[i].Equals("BUZZER1") ||
                                                        cmd.Parameter[i].Equals("BUZZER2")))
                                                    {
                                                        point = cmd.Parameter[i];
                                                    }
                                                    else if (type.Equals("PORT") &&
                                                       (cmd.Parameter[i].Equals("LOAD") ||
                                                       cmd.Parameter[i].Equals("UNLOAD") ||
                                                       cmd.Parameter[i].Equals("ACCESS")))
                                                    {
                                                        point = cmd.Parameter[i];
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
                                                        if ((point.Equals("BUZZER1") || point.Equals("BUZZER2")) && cmd.Parameter[i].Equals("BLINK"))
                                                        {
                                                            //命令錯誤
                                                            SendNak(WaitForHandle, "Command format error.");
                                                            SendInfo(WaitForHandle);
                                                            return;
                                                        }
                                                        else
                                                        {
                                                            state = cmd.Parameter[i].Replace("ON", "TRUE").Replace("OFF", "FALSE");
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
                                        ErrorMessage = "";

                                        if (type.Equals("PORT"))
                                        {
                                            Node port = NodeManagement.Get(Target);
                                            switch (point)
                                            {
                                                case "LOAD":
                                                    TaskName = "SET_LOAD_INDICATOR";
                                                    port.Load_LED = state;
                                                    switch (state)
                                                    {
                                                        case "TRUE":
                                                            state = "1";
                                                            break;
                                                        case "FALSE":
                                                            state = "0";
                                                            break;
                                                        case "BLINK":
                                                            state = "2";
                                                            break;
                                                    }
                                                    break;
                                                case "UNLOAD":
                                                    TaskName = "SET_UNLOAD_INDICATOR";
                                                    port.UnLoad_LED = state;
                                                    switch (state)
                                                    {
                                                        case "TRUE":
                                                            state = "1";
                                                            break;
                                                        case "FALSE":
                                                            state = "0";
                                                            break;
                                                        case "BLINK":
                                                            state = "2";
                                                            break;
                                                    }
                                                    break;
                                                case "ACCESS":
                                                    TaskName = "SET_OPACCESS_INDICATOR";
                                                    port.AccessSW_LED = state;
                                                    switch (state)
                                                    {
                                                        case "TRUE":
                                                            state = "1";
                                                            break;
                                                        case "FALSE":
                                                            state = "0";
                                                            break;
                                                        case "BLINK":
                                                            state = "2";
                                                            break;
                                                    }
                                                    break;
                                            }

                                            Dictionary<string, string> param = new Dictionary<string, string>();
                                            param.Add("@Target", Target);
                                            param.Add("@state", state);
                                            TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                            if (!ErrorMessage.Equals(""))
                                            {
                                                SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                               // SendInfo(WaitForHandle);
                                            }
                                            else
                                            {
                                                SendAck(WaitForHandle);

                                            }
                                           
                                        }
                                        else if (type.Equals("STOWER"))
                                        {
                                            switch (point)
                                            {
                                                case "RED":
                                                    switch (state)
                                                    {
                                                        case "TRUE":
                                                            RouteControl.DIO.SetIO("RED", "True");
                                                            break;
                                                        case "FALSE":
                                                            RouteControl.DIO.SetIO("RED", "False");
                                                            break;
                                                        case "BLINK":
                                                            RouteControl.DIO.SetBlink("RED", "True");
                                                            break;
                                                    }
                                                    break;
                                                case "YELLOW":
                                                    switch (state)
                                                    {
                                                        case "TRUE":
                                                            RouteControl.DIO.SetIO("ORANGE", "True");
                                                            break;
                                                        case "FALSE":
                                                            RouteControl.DIO.SetIO("ORANGE", "False");
                                                            break;
                                                        case "BLINK":
                                                            RouteControl.DIO.SetBlink("ORANGE", "True");
                                                            break;
                                                    }
                                                    break;
                                                case "GREEN":
                                                    switch (state)
                                                    {
                                                        case "TRUE":
                                                            RouteControl.DIO.SetIO("GREEN", "True");
                                                            break;
                                                        case "FALSE":
                                                            RouteControl.DIO.SetIO("GREEN", "False");
                                                            break;
                                                        case "BLINK":
                                                            RouteControl.DIO.SetBlink("GREEN", "True");
                                                            break;
                                                    }
                                                    break;
                                                case "BLUE":
                                                    switch (state)
                                                    {
                                                        case "TRUE":
                                                            RouteControl.DIO.SetIO("BLUE", "True");
                                                            break;
                                                        case "FALSE":
                                                            RouteControl.DIO.SetIO("BLUE", "False");
                                                            break;
                                                        case "BLINK":
                                                            RouteControl.DIO.SetBlink("BLUE", "True");
                                                            break;
                                                    }
                                                    break;
                                                case "BUZZER1":
                                                    switch (state)
                                                    {
                                                        case "TRUE":
                                                            RouteControl.DIO.SetIO("BUZZER1", "True");
                                                            break;
                                                        case "FALSE":
                                                            RouteControl.DIO.SetIO("BUZZER1", "False");
                                                            break;
                                                    }
                                                    break;
                                                case "BUZZER2":
                                                    switch (state)
                                                    {
                                                        case "TRUE":
                                                            RouteControl.DIO.SetIO("BUZZER2", "True");
                                                            break;
                                                        case "FALSE":
                                                            RouteControl.DIO.SetIO("BUZZER2", "False");
                                                            break;
                                                    }
                                                    break;
                                            }
                                            SendAck(WaitForHandle);
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
                                case "EVENT":
                                    try
                                    {
                                        //檢查命令格式
                                        string EventName = "";
                                        string State = "";
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
                                                        EventName = cmd.Parameter[i];
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
                                                        State = cmd.Parameter[i];
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
                                        //if (State.Equals("ON"))
                                        //{
                                        //    EventHandling = WaitForHandle;
                                        //}
                                        switch (EventName)
                                        {
                                            case "ALL":
                                                if (State.Equals("ON"))
                                                {
                                                    Events.FFU = true;
                                                    Events.MAPDT = true;
                                                    Events.PORT = true;
                                                    Events.PRS = true;
                                                    Events.SYSTEM = true;
                                                    Events.TRANSREQ = true;
                                                }
                                                else if (State.Equals("OFF"))
                                                {
                                                    Events.FFU = false;
                                                    Events.MAPDT = false;
                                                    Events.PORT = false;
                                                    Events.PRS = false;
                                                    Events.SYSTEM = false;
                                                    Events.TRANSREQ = false;
                                                }
                                                break;
                                            case "MAPDT":
                                                if (State.Equals("ON"))
                                                {
                                                    Events.MAPDT = true;
                                                }
                                                else if (State.Equals("OFF"))
                                                {
                                                    Events.MAPDT = false;
                                                }
                                                break;
                                            case "TRANSREQ":
                                                if (State.Equals("ON"))
                                                {
                                                    Events.TRANSREQ = true;
                                                }
                                                else if (State.Equals("OFF"))
                                                {
                                                    Events.TRANSREQ = false;
                                                }
                                                break;
                                            case "SYSTEM":
                                                if (State.Equals("ON"))
                                                {
                                                    Events.SYSTEM = true;
                                                }
                                                else if (State.Equals("OFF"))
                                                {
                                                    Events.SYSTEM = false;
                                                }
                                                break;
                                            case "PORT":
                                                if (State.Equals("ON"))
                                                {
                                                    Events.PORT = true;
                                                }
                                                else if (State.Equals("OFF"))
                                                {
                                                    Events.PORT = false;
                                                }
                                                break;
                                            case "PRS":
                                                if (State.Equals("ON"))
                                                {
                                                    Events.PRS = true;
                                                }
                                                else if (State.Equals("OFF"))
                                                {
                                                    Events.PRS = false;
                                                }
                                                break;
                                            case "FFU":
                                                if (State.Equals("ON"))
                                                {
                                                    Events.FFU = true;
                                                }
                                                else if (State.Equals("OFF"))
                                                {
                                                    Events.FFU = false;
                                                }
                                                break;

                                        }
                                        Events.Save();
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
                                case "SIZE":
                                    try
                                    {
                                        //int no = 0;
                                        //檢查命令格式
                                        Target = "";
                                        string Size = "";
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
                                                        Target = "ROBOT01";
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("P") != -1 &&
                                                        int.TryParse(cmd.Parameter[i].Replace("P", ""), out no) &&
                                                        cmd.Parameter[i].Replace("P", "").Length == 3)
                                                    {
                                                        Target = NodeNameConvert(cmd.Parameter[i].Substring(0, 2), "LOADPORT");
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
                                                    else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                        int.TryParse(cmd.Parameter[i].Replace("LL", ""), out no) &&
                                                        (cmd.Parameter[i].Replace("LL", "").Length == 3))
                                                    {
                                                        Target = NodeNameConvert(cmd.Parameter[i].Substring(0, 2), "STAGE");
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
                                                        Size = cmd.Parameter[i];
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
                                        Node TargetNode = NodeManagement.Get(Target);
                                        if (TargetNode != null)
                                        {
                                            TargetNode.WaferSize = Size + "MM";
                                            SendAck(WaitForHandle);
                                            SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", Target + " not found");
                                           // SendInfo(WaitForHandle);
                                        }


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

                                        TaskName = "";
                                        Target = "";
                                        for (int i = 0; i < cmd.Parameter.Count; i++)
                                        {
                                            switch (i)
                                            {
                                                case 0:
                                                    TaskName = cmd.Parameter[i];
                                                    //If the parameter is omitted, the system acts in the same manner as when "ALL" is designated.
                                                    if (cmd.Parameter[i].Equals("ALL"))
                                                    {
                                                        TaskName = "ALL";
                                                        Target = "ALL";
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

                                        ErrorMessage = "";
                                        TaskName += "_Init";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        WaitForHandle.Cmd.Target = Target;
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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
                                        TaskName = "";
                                        Target = "";
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
                                                        Target = "ALL";
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

                                        ErrorMessage = "";
                                        TaskName += "_ORGSH";

                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        WaitForHandle.Cmd.Target = Target;

                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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

                                        TaskName = "";
                                        Target = "";
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

                                        ErrorMessage = "";
                                        TaskName = "LOCK";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);

                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            SendAck(WaitForHandle);
                                            Node port = NodeManagement.Get(Target);
                                            port.Foup_Lock = true;
                                            On_Event_Trigger("SIGSTAT", "PORT", Target, "ALL");
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
                                        TaskName = "";
                                        Target = "";
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

                                        ErrorMessage = "";
                                        TaskName = "UNLOCK";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            SendAck(WaitForHandle);
                                            Node port = NodeManagement.Get(Target);
                                            port.Foup_Lock = false;
                                            On_Event_Trigger("SIGSTAT", "PORT", Target, "ALL");
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
                                        TaskName = "";
                                        Target = "";

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

                                        ErrorMessage = "";
                                        TaskName = "DOCK";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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
                                        TaskName = "";
                                        Target = "";
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

                                        ErrorMessage = "";
                                        TaskName = "UNDOCK";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                          //  SendInfo(WaitForHandle);
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
                                        TaskName = "";
                                        Target = "";
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

                                        ErrorMessage = "";
                                        TaskName = "OPEN";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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
                                        TaskName = "";
                                        Target = "";
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

                                        ErrorMessage = "";
                                        TaskName = "CLOSE";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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
                                        TaskName = "";
                                        Target = "";
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

                                        ErrorMessage = "";
                                        TaskName = "WAFSH";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", Target);
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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
                                        TaskName = "";
                                        string Slot = "";
                                        Target = "";
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
                                                    else if (cmd.Parameter[i].IndexOf("BF1") != -1)
                                                    {
                                                        //Target = cmd.Parameter[i].Substring(0, 2);

                                                        if (!int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no))
                                                        {
                                                            Slot = "1";
                                                        }
                                                        else
                                                        {
                                                            Slot = no.ToString();
                                                        }

                                                        Position = "ALIGNER02";
                                                        //TargetCheckMethod = "ReadStatus";
                                                    }
                                                    //else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    //    (cmd.Parameter[i].Replace("LL", "").Length == 1 ||
                                                    //    (cmd.Parameter[i].Replace("LL", "").Length == 3) && int.TryParse(cmd.Parameter[i].Substring(3), out no)))
                                                    //{
                                                    //    Target = cmd.Parameter[i].Substring(0, 3);
                                                    //    Position = NodeNameConvert(Target, "STAGE");
                                                    //    if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    //    {
                                                    //        Slot = no.ToString();
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        Slot = "1";
                                                    //    }
                                                    //}
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

                                        ErrorMessage = "";
                                        TaskName += "_GOTO";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Slot", Slot);
                                        param.Add("@Arm", Arm);
                                        param.Add("@Position", Position);
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                            //SendInfo(WaitForHandle);
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
                                        TaskName = "";
                                        string Slot = "";
                                        Target = "";
                                        string Position = "";
                                        string Arm = "";
                                        string TargetCheckMethod = "";

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
                                                        TargetCheckMethod = "ReadStatus";
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                        int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no))
                                                    {
                                                        Target = cmd.Parameter[i];
                                                        Slot = "1";
                                                        Position = NodeNameConvert(Target, "ALIGNER");
                                                        TargetCheckMethod = "GetRIO";
                                                    }
                                                    else if (cmd.Parameter[i].Equals("ALIGN"))
                                                    {
                                                        Target = "ALIGN1";
                                                        Slot = "1";
                                                        Position = NodeNameConvert(Target, "ALIGNER");
                                                        TargetCheckMethod = "GetRIO";
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("BF1") != -1)
                                                    {
                                                        //Target = cmd.Parameter[i].Substring(0, 2);

                                                        if (!int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no))
                                                        {
                                                            Slot = "1";
                                                        }
                                                        else
                                                        {
                                                            Slot = no.ToString();
                                                        }

                                                        Position = "LOADLOCK01";
                                                        TargetCheckMethod = "ReadStatus";
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("BF2") != -1)
                                                    {
                                                        //Target = cmd.Parameter[i].Substring(0, 2);

                                                        if (!int.TryParse(cmd.Parameter[i].Replace("BF2", ""), out no))
                                                        {
                                                            Slot = "1";
                                                        }
                                                        else
                                                        {
                                                            Slot = no.ToString();
                                                        }

                                                        Position = "LOADLOCK02";
                                                        TargetCheckMethod = "ReadStatus";
                                                    }
                                                    //else if (cmd.Parameter[i].Equals("LLA"))
                                                    //{
                                                    //    Slot = no.ToString().Substring(1);
                                                    //    Position = "LOADPORT03";
                                                    //    TargetCheckMethod = "ReadStatus";
                                                    //}
                                                    //else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    //    (cmd.Parameter[i].Replace("LL", "").Length == 1 ||
                                                    //    (cmd.Parameter[i].Replace("LL", "").Length == 3 && int.TryParse(cmd.Parameter[i].Substring(3), out no))))
                                                    //{
                                                    //    Target = cmd.Parameter[i].Substring(0, 3);
                                                    //    Position = NodeNameConvert(Target, "STAGE");
                                                    //    if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    //    {
                                                    //        Slot = no.ToString();
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        Slot = "1";
                                                    //    }
                                                    //    TargetCheckMethod = "GetRIO";
                                                    //}
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
                                        //Node n = NodeManagement.Get(Position);

                                        //if (n.Type.ToUpper().Equals("LOADPORT") && !SaftyCheckByPass)
                                        //{
                                        //    //檢查Slot是否安全
                                        //    if (!n.IsMapping)
                                        //    {
                                        //        SendCancel(WaitForHandle, "SAFTY", "ARM1", ErrorMessage);
                                        //       // SendInfo(WaitForHandle);
                                        //        return;
                                        //    }
                                        //    else
                                        //    {
                                        //        int slotNo = 0;
                                        //        if (int.TryParse(Slot, out slotNo))
                                        //        {
                                        //            Job SlotData = null;
                                        //            n.JobList.TryGetValue(slotNo.ToString(), out SlotData);
                                        //            if (SlotData.MapFlag && !SlotData.ErrPosition)
                                        //            {

                                        //            }
                                        //            else
                                        //            {
                                        //                SendCancel(WaitForHandle, "SAFTY", "ARM1", ErrorMessage);
                                        //               // SendInfo(WaitForHandle);
                                        //                return;
                                        //            }

                                        //        }
                                        //        else
                                        //        {
                                        //            SendCancel(WaitForHandle, "SAFTY", "ARM1", ErrorMessage);
                                        //           // SendInfo(WaitForHandle);
                                        //            return;
                                        //        }
                                        //    }

                                        //}

                                        ErrorMessage = "";
                                        TaskName = "LOAD";
                                        if (!SaftyCheckByPass)
                                        {
                                            TaskName += "_SafetyCheck";
                                        }
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", "ROBOT01");
                                        param.Add("@Slot", Slot);
                                        param.Add("@Arm", Arm);
                                        param.Add("@Position", Position);
                                        param.Add("@Method", TargetCheckMethod);



                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            //SendAck(WaitForHandle);
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
                                        TaskName = "";
                                        string Slot = "";
                                        Target = "";
                                        string Position = "";
                                        string Arm = "";
                                        string TargetCheckMethod = "";
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
                                                        TargetCheckMethod = "ReadStatus";
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("ALIGN") != -1 &&
                                                        int.TryParse(cmd.Parameter[i].Replace("ALIGN", ""), out no))
                                                    {
                                                        Target = cmd.Parameter[i];
                                                        Slot = "1";
                                                        Position = NodeNameConvert(Target, "ALIGNER");
                                                        TargetCheckMethod = "GetRIO";
                                                    }
                                                    else if (cmd.Parameter[i].Equals("ALIGN"))
                                                    {
                                                        Target = "ALIGN1";
                                                        Slot = "1";
                                                        Position = NodeNameConvert(Target, "ALIGNER");
                                                        TargetCheckMethod = "GetRIO";
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("BF1") != -1)
                                                    {
                                                        //Target = cmd.Parameter[i].Substring(0, 2);

                                                        if (!int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no))
                                                        {
                                                            Slot = "1";
                                                        }
                                                        else
                                                        {
                                                            Slot = no.ToString();
                                                        }

                                                        Position = "LOADLOCK01";
                                                        TargetCheckMethod = "ReadStatus";
                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("BF2") != -1)
                                                    {
                                                        //Target = cmd.Parameter[i].Substring(0, 2);

                                                        if (!int.TryParse(cmd.Parameter[i].Replace("BF2", ""), out no))
                                                        {
                                                            Slot = "1";
                                                        }
                                                        else
                                                        {
                                                            Slot = no.ToString();
                                                        }

                                                        Position = "LOADLOCK02";
                                                        TargetCheckMethod = "ReadStatus";
                                                    }
                                                    //else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    //    (cmd.Parameter[i].Replace("LL", "").Length == 1 ||
                                                    //    (cmd.Parameter[i].Replace("LL", "").Length == 3 && int.TryParse(cmd.Parameter[i].Substring(3), out no))))
                                                    //{
                                                    //    Target = cmd.Parameter[i].Substring(0, 3);
                                                    //    Position = NodeNameConvert(Target, "STAGE");
                                                    //    if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    //    {
                                                    //        Slot = no.ToString();
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        Slot = "1";
                                                    //    }
                                                    //    TargetCheckMethod = "GetRIO";
                                                    //}
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
                                        Node n = NodeManagement.Get(Position);

                                        if (n.Type.ToUpper().Equals("LOADPORT") && !SaftyCheckByPass)
                                        {
                                            //檢查Slot是否安全
                                            if (n.MappingResult.Equals(""))
                                            {
                                                SendCancel(WaitForHandle, "SAFTY", Target, ErrorMessage);
                                              //  SendInfo(WaitForHandle);
                                                return;
                                            }
                                            else
                                            {
                                                int slotNo = 0;
                                                if (int.TryParse(Slot, out slotNo))
                                                {
                                                    Job SlotData = null;
                                                    n.JobList.TryGetValue(slotNo.ToString(), out SlotData);
                                                    if (!SlotData.MapFlag)
                                                    {

                                                    }
                                                    else
                                                    {
                                                        SendCancel(WaitForHandle, "SAFTY", Target, ErrorMessage);
                                                       // SendInfo(WaitForHandle);
                                                        return;
                                                    }

                                                }
                                                else
                                                {
                                                    SendCancel(WaitForHandle, "SAFTY", Target, ErrorMessage);
                                                  //  SendInfo(WaitForHandle);
                                                    return;
                                                }
                                            }

                                        }
                                        ErrorMessage = "";
                                        TaskName = "UNLOAD";
                                        if (!SaftyCheckByPass)
                                        {
                                            TaskName += "_SafetyCheck";
                                        }
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", "ROBOT01");
                                        param.Add("@Slot", Slot);
                                        param.Add("@Arm", Arm);
                                        param.Add("@Position", Position);
                                        param.Add("@Method", TargetCheckMethod);
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            //SendAck(WaitForHandle);
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
                                        TaskName = "";
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
                                                    else if (cmd.Parameter[i].IndexOf("BF1") != -1)
                                                    {

                                                        if (i == 0)
                                                        {
                                                            FromTarget = "LOADLOCK01";
                                                            //int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no);
                                                            if (!int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no))
                                                            {
                                                                FromSlot = "1";
                                                            }
                                                            else
                                                            {
                                                                FromSlot = no.ToString();
                                                            }


                                                        }
                                                        else if (i == 3)
                                                        {
                                                            ToTarget = "LOADLOCK01";
                                                            //int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no);
                                                            if (!int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no))
                                                            {
                                                                ToSlot = "1";
                                                            }
                                                            else
                                                            {
                                                                ToSlot = no.ToString();
                                                            }
                                                            ToSlot = "1";
                                                        }

                                                    }
                                                    else if (cmd.Parameter[i].IndexOf("BF2") != -1)
                                                    {

                                                        if (i == 0)
                                                        {
                                                            FromTarget = "LOADLOCK02";
                                                            //int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no);
                                                            if (!int.TryParse(cmd.Parameter[i].Replace("BF2", ""), out no))
                                                            {
                                                                FromSlot = "1";
                                                            }
                                                            else
                                                            {
                                                                FromSlot = no.ToString();
                                                            }


                                                        }
                                                        else if (i == 3)
                                                        {
                                                            ToTarget = "LOADLOCK02";
                                                            //int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no);
                                                            if (!int.TryParse(cmd.Parameter[i].Replace("BF2", ""), out no))
                                                            {
                                                                ToSlot = "1";
                                                            }
                                                            else
                                                            {
                                                                ToSlot = no.ToString();
                                                            }
                                                            ToSlot = "1";
                                                        }

                                                    }
                                                    //else if (cmd.Parameter[i].IndexOf("LL") != -1 &&
                                                    //    (cmd.Parameter[i].Replace("LL", "").Length == 1 ||
                                                    //    (cmd.Parameter[i].Replace("LL", "").Length == 3 && int.TryParse(cmd.Parameter[i].Substring(3), out no))))
                                                    //{
                                                    //    if (i == 0)
                                                    //    {
                                                    //        if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    //        {
                                                    //            FromTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                    //            FromSlot = no.ToString();
                                                    //        }
                                                    //        else
                                                    //        {
                                                    //            FromTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                    //            FromSlot = "1";
                                                    //        }

                                                    //    }
                                                    //    else if (i == 3)
                                                    //    {
                                                    //        if (cmd.Parameter[i].Replace("LL", "").Length == 3)
                                                    //        {
                                                    //            ToTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                    //            ToSlot = no.ToString();
                                                    //        }
                                                    //        else
                                                    //        {
                                                    //            ToTarget = NodeNameConvert(cmd.Parameter[i].Substring(0, 3), "STAGE");
                                                    //            ToSlot = "1";
                                                    //        }
                                                    //    }
                                                    //}
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
                                        Node n = NodeManagement.Get(FromTarget);

                                        if (n.Type.ToUpper().Equals("LOADPORT") && !SaftyCheckByPass)
                                        {
                                            //檢查Slot是否安全
                                            if (n.MappingResult.Equals(""))
                                            {
                                                SendCancel(WaitForHandle, "SAFTY", Target, ErrorMessage);
                                               // SendInfo(WaitForHandle);
                                                return;
                                            }
                                            else
                                            {
                                                int slotNo = 0;
                                                if (int.TryParse(FromSlot, out slotNo))
                                                {
                                                    Job SlotData = null;
                                                    n.JobList.TryGetValue(slotNo.ToString(), out SlotData);
                                                    if (SlotData.MapFlag && !SlotData.ErrPosition)
                                                    {

                                                    }
                                                    else
                                                    {
                                                        SendCancel(WaitForHandle, "SAFTY", Target, ErrorMessage);
                                                       // SendInfo(WaitForHandle);
                                                        return;
                                                    }
                                                }
                                                else
                                                {
                                                    SendCancel(WaitForHandle, "SAFTY", Target, ErrorMessage);
                                                   // SendInfo(WaitForHandle);
                                                    return;
                                                }
                                            }

                                        }

                                        n = NodeManagement.Get(ToTarget);

                                        if (n.Type.ToUpper().Equals("LOADPORT") && !SaftyCheckByPass)
                                        {
                                            //檢查Slot是否安全
                                            if (n.MappingResult.Equals(""))
                                            {
                                                SendCancel(WaitForHandle, "SAFTY", Target, ErrorMessage);
                                               // SendInfo(WaitForHandle);
                                                return;
                                            }
                                            else
                                            {
                                                int slotNo = 0;
                                                if (int.TryParse(ToSlot, out slotNo))
                                                {
                                                    Job SlotData = null;
                                                    n.JobList.TryGetValue(slotNo.ToString(), out SlotData);
                                                    if (!SlotData.MapFlag)
                                                    {

                                                    }
                                                    else
                                                    {
                                                        SendCancel(WaitForHandle, "SAFTY", Target, ErrorMessage);
                                                       // SendInfo(WaitForHandle);
                                                        return;
                                                    }
                                                }
                                                else
                                                {
                                                    SendCancel(WaitForHandle, "SAFTY", Target, ErrorMessage);
                                                   // SendInfo(WaitForHandle);
                                                    return;
                                                }
                                            }

                                        }

                                        ErrorMessage = "";
                                        TaskName = "TRANS";
                                        if (!SaftyCheckByPass)
                                        {
                                            TaskName += "_SafetyCheck";
                                        }
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@FromPosition", FromTarget);
                                        param.Add("@ToPosition", ToTarget);
                                        param.Add("@FromARM", FromARM);
                                        param.Add("@ToARM", ToARM);
                                        param.Add("@FromSlot", FromSlot);
                                        param.Add("@ToSlot", ToSlot);
                                        param.Add("@Target", "ROBOT01");
                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            //SendAck(WaitForHandle);
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
                                        TaskName = "";
                                        Target = "";
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
                                                    else if (cmd.Parameter[i].IndexOf("BF1") != -1)
                                                    {
                                                        //Target = cmd.Parameter[i].Substring(0, 2);
                                                        //int.TryParse(cmd.Parameter[i].Replace("BF01", ""), out no);
                                                        if (!int.TryParse(cmd.Parameter[i].Replace("BF1", ""), out no))
                                                        {
                                                            TargetSlot = "1";
                                                        }
                                                        else
                                                        {
                                                            TargetSlot = no.ToString();
                                                        }

                                                        Target = "ALIGNER02";
                                                        //TargetCheckMethod = "ReadStatus";
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

                                        ErrorMessage = "";
                                        TaskName = "CHANGE";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Self", "ROBOT01");
                                        param.Add("@Position", Target);
                                        param.Add("@Slot", TargetSlot);
                                        param.Add("@CarryOutARM", CarryOutARM);
                                        param.Add("@CarryInARM", CarryInARM);

                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
                                        }
                                        else
                                        {
                                            //SendAck(WaitForHandle);
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
                                        TaskName = "";
                                        string Slot = "";
                                        Target = "";
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

                                        ErrorMessage = "";
                                        TaskName = "ALIGN";
                                        Dictionary<string, string> param = new Dictionary<string, string>();
                                        param.Add("@Target", NodeNameConvert(Target, "ALIGNER"));

                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName, param);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                            //SendInfo(WaitForHandle);
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
                                        TaskName = "";
                                        Target = "";
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
                                        ErrorMessage = "";
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
                                           // SendInfo(WaitForHandle);
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

                                        ErrorMessage = "";
                                        TaskName = "HOLD";
                                        //Dictionary<string, string> param = new Dictionary<string, string>();
                                        if (!NodeManagement.Get("ROBOT01").IsExcuting)
                                        {
                                            SendCancel(WaitForHandle, "NOTINMOTION", "ROBOT", ErrorMessage);
                                           // SendInfo(WaitForHandle);
                                            return;
                                        }

                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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

                                        ErrorMessage = "";
                                        TaskName = "RESTR";
                                        //Dictionary<string, string> param = new Dictionary<string, string>();


                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                            //SendInfo(WaitForHandle);
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

                                        ErrorMessage = "";
                                        TaskName = "ABORT";
                                        //Dictionary<string, string> param = new Dictionary<string, string>();


                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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

                                        ErrorMessage = "";
                                        TaskName = "EMS";
                                        //Dictionary<string, string> param = new Dictionary<string, string>();


                                        TaskJobManagment.Excute(WaitForHandle.ID, out ErrorMessage, TaskName);

                                        if (!ErrorMessage.Equals(""))
                                        {
                                            SendCancel(WaitForHandle, ErrorCategory.CancelFactor.NOLINK, "", ErrorMessage);
                                           // SendInfo(WaitForHandle);
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
            if (TimeOutCmd.INF_RetryCount < 0)
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

        public void On_TaskJob_Ack(string TaskID)
        {
            OnHandling WaitForHandle;

            if (OnHandlingCmds.TryGetValue(TaskID, out WaitForHandle))
            {
                SendAck(WaitForHandle);
            }
        }

        public void On_TaskJob_Finished(string TaskID)
        {
            OnHandling WaitForHandle;
            Node Target = null;
            string Data1 = "";
            if (OnHandlingCmds.TryGetValue(TaskID, out WaitForHandle))
            {
                switch (WaitForHandle.Cmd.CommandType)
                {
                    case "MOV":
                        switch (WaitForHandle.Cmd.Command)
                        {
                            case "INIT":
                                if (WaitForHandle.Cmd.Target.Equals("ALL"))
                                {
                                    EFEM_State = NodeManagement.Get("ROBOT01").State;
                                }
                                break;
                            case "ORGSH":
                                if (WaitForHandle.Cmd.Target.Equals("ALL")|| WaitForHandle.Cmd.Target.IndexOf("ROBOT")!=-1)
                                {
                                    EFEM_State = "Ready";
                                }
                                break;
                        }
                        SendInfo(WaitForHandle);
                        break;
                    case "GET":
                        switch (WaitForHandle.Cmd.Command)
                        {
                            case "MAPDT":
                                Target = NodeManagement.Get(WaitForHandle.Cmd.Target);
                                string Mapping = Target.MappingResult;
                                Mapping.Replace("2", "3").Replace("E", "3").Replace("W", "7").Replace("?", "9");


                                SendInfo(WaitForHandle, Mapping, "");
                                break;
                            case "CLAMP":
                                Target = NodeManagement.Get(WaitForHandle.Cmd.Target);

                                switch (WaitForHandle.Cmd.Arm)
                                {
                                    case "1":
                                        if (Target.RArmClamp&& Target.RArmUnClamp)
                                        {
                                            Data1 = "ON";
                                        }
                                        else
                                        {
                                            Data1 = "OFF";
                                        }
                                        break;
                                    case "2":
                                        if (Target.LArmClamp && Target.LArmUnClamp)
                                        {
                                            Data1 = "ON";
                                        }
                                        else
                                        {
                                            Data1 = "OFF";
                                        }
                                        break;
                                }
                                SendInfo(WaitForHandle, Data1, "");
                                break;
                            case "CSTID":
                                Target = NodeManagement.Get(WaitForHandle.Cmd.Target);

                                SendInfo(WaitForHandle, Target.FoupID, "");
                                break;
                        }
                        break;
                    case "SET":
                        switch (WaitForHandle.Cmd.Command)
                        {
                            case "SIGOUT":
                                string[] CmdAry = WaitForHandle.Cmd.OrgMsg.Replace(";\r", "").Split(new char[] { ':', '/' });
                                if (!CmdAry[2].Equals("SYSTEM"))
                                {
                                    SendInfo(WaitForHandle);
                                    On_Event_Trigger("SIGSTAT", "PORT", NodeNameConvert(CmdAry[2], "LOADPORT"), "ALL");
                                    return;
                                }
                                break;
                        }
                        SendInfo(WaitForHandle);
                        break;
                    default:
                        SendInfo(WaitForHandle);
                        break;
                }


            }
            else
            {
                logger.Error("On_TaskJob_Aborted 找不到 TaskID:" + TaskID);
            }
        }

        public void On_TaskJob_Aborted(string TaskID, string NodeName, string ReportType, string Message)
        {
            OnHandling WaitForHandle;
            if (OnHandlingCmds.TryGetValue(TaskID, out WaitForHandle))
            {
                try
                {
                    AlarmMessage alm = AlmMapping.Get(NodeName, Message);


                    if (ReportType.Equals("ABS"))
                    {
                        //var findx = from Handling in OnHandlingCmds.Values.ToList()
                        //            where Handling.Cmd.Command.Equals("ABORT") && Handling.Cmd.CommandType.Equals(CommandType.MOV)
                        //            select Handling;
                        //if (findx.Count() != 0 && Message.Equals("84805000"))
                        //{

                        //    OnHandlingCmds.TryRemove(WaitForHandle.ID,out WaitForHandle);
                        //    return;
                        //}

                        LastError = alm;
                        if (LastError.Position.Equals(""))
                        {
                            if (NodeName.ToUpper().IndexOf("LOADPORT") != -1)
                            {
                                alm.Position = "P" + NodeName.Replace("LOADPORT0", "");
                            }

                        }
                        SendABS(WaitForHandle, LastError.Code_Group, LastError.Position);
                    }
                    else if (ReportType.Equals("CAN"))
                    {
                        SendCancel(WaitForHandle, alm.Code_Group, alm.Position, "");
                       // SendInfo(WaitForHandle);
                    }

                }
                catch (Exception e)
                {
                    logger.Error("(GetAlarmMessage)" + e.Message + "\n" + e.StackTrace);
                    SendABS(WaitForHandle, "TEST", Message);
                }

            }
            else
            {
                logger.Error("On_TaskJob_Aborted 找不到 TaskID:" + TaskID + " Message:" + Message);
            }
        }
        
        public void On_Event_Trigger(string Type, string Source, string Name, string Value)
        {
            try
            {
                switch (Type)
                {
                    case "MAPDT":
                        if (Events.MAPDT)
                        {
                            int pNo = 0;
                            int.TryParse(Name.Replace("LOADPORT", ""), out pNo);
                            Value = Value.Replace("2", "3").Replace("E", "3").Replace("W", "7").Replace("?", "9");
                            SendEvent(EventHandling, "MAPDT", "P" + pNo, Value, "");
                        }
                        break;
                    case "SIGSTAT":
                        switch (Source)
                        {
                            case "SYSTEM":
                                string Data1 = "00000000000000000000000000000000";
                                string Data2 = "00000000000000000000000000000000";
                                string[] key_pairs = Value.Split(',');
                                foreach (string key_pair_str in key_pairs)
                                {
                                    string[] key_pair = key_pair_str.Split('=');
                                    string key = key_pair[0];
                                    string val = key_pair[1];
                                    switch (key)
                                    {
                                        case "VACUUM":
                                            Data1 = Data1.Remove(0, 1);
                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data1 = Data1.Insert(0, "1");
                                            }
                                            else
                                            {
                                                Data1 = Data1.Insert(0, "0");
                                            }
                                            break;
                                        case "AIR":
                                            Data1 = Data1.Remove(2, 1);
                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data1 = Data1.Insert(2, "1");
                                            }
                                            else
                                            {
                                                Data1 = Data1.Insert(2, "0");
                                            }
                                            break;
                                        case "DEFFERENTIALPRESSUREALARM1":
                                            Data1 = Data1.Remove(4, 1);
                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data1 = Data1.Insert(4, "1");
                                            }
                                            else
                                            {
                                                Data1 = Data1.Insert(4, "0");
                                            }
                                            break;
                                        case "DEFFERENTIALPRESSUREALARM2":
                                            Data1 = Data1.Remove(5, 1);
                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data1 = Data1.Insert(5, "1");
                                            }
                                            else
                                            {
                                                Data1 = Data1.Insert(5, "0");
                                            }
                                            break;
                                        case "FFU":
                                            Data1 = Data1.Remove(6, 1);
                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data1 = Data1.Insert(6, "1");
                                            }
                                            else
                                            {
                                                Data1 = Data1.Insert(6, "0");
                                            }
                                            break;
                                        case "IONIZERALARM":
                                            Data1 = Data1.Remove(7, 1);
                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data1 = Data1.Insert(7, "1");
                                            }
                                            else
                                            {
                                                Data1 = Data1.Insert(7, "0");
                                            }
                                            break;
                                        case "DOORSWITCH":
                                            Data1 = Data1.Remove(10, 1);
                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data1 = Data1.Insert(10, "1");
                                            }
                                            else
                                            {
                                                Data1 = Data1.Insert(10, "0");
                                            }
                                            break;
                                        case "SAFETYRELAY":
                                            Data1 = Data1.Remove(11, 1);
                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data1 = Data1.Insert(11, "1");
                                            }
                                            else
                                            {
                                                Data1 = Data1.Insert(11, "0");
                                            }
                                            break;
                                        case "RED":

                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data2 = Data2.Remove(0, 1);
                                                Data2 = Data2.Insert(0, "1");
                                            }
                                            else if (val.ToUpper().Equals("BLINK"))
                                            {
                                                Data2 = Data2.Remove(5, 1);
                                                Data2 = Data2.Insert(5, "1");
                                            }
                                            else
                                            {
                                                Data2 = Data2.Remove(0, 1);
                                                Data2 = Data2.Insert(0, "0");
                                                Data2 = Data2.Remove(5, 1);
                                                Data2 = Data2.Insert(5, "0");
                                            }
                                            break;
                                        case "ORANGE":

                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data2 = Data2.Remove(1, 1);
                                                Data2 = Data2.Insert(1, "1");
                                            }
                                            else if (val.ToUpper().Equals("BLINK"))
                                            {
                                                Data2 = Data2.Remove(6, 1);
                                                Data2 = Data2.Insert(6, "1");
                                            }
                                            else
                                            {
                                                Data2 = Data2.Remove(1, 1);
                                                Data2 = Data2.Insert(1, "0");
                                                Data2 = Data2.Remove(6, 1);
                                                Data2 = Data2.Insert(6, "0");
                                            }
                                            break;
                                        case "GREEN":

                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data2 = Data2.Remove(2, 1);
                                                Data2 = Data2.Insert(2, "1");
                                            }
                                            else if (val.ToUpper().Equals("BLINK"))
                                            {
                                                Data2 = Data2.Remove(7, 1);
                                                Data2 = Data2.Insert(7, "1");
                                            }
                                            else
                                            {
                                                Data2 = Data2.Remove(2, 1);
                                                Data2 = Data2.Insert(2, "0");
                                                Data2 = Data2.Remove(7, 1);
                                                Data2 = Data2.Insert(7, "0");
                                            }
                                            break;
                                        case "BLUE":

                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data2 = Data2.Remove(3, 1);
                                                Data2 = Data2.Insert(3, "1");
                                            }
                                            else if (val.ToUpper().Equals("BLINK"))
                                            {
                                                Data2 = Data2.Remove(8, 1);
                                                Data2 = Data2.Insert(8, "1");
                                            }
                                            else
                                            {
                                                Data2 = Data2.Remove(3, 1);
                                                Data2 = Data2.Insert(3, "0");
                                                Data2 = Data2.Remove(8, 1);
                                                Data2 = Data2.Insert(8, "0");
                                            }
                                            break;
                                        case "BUZZER1":

                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data2 = Data2.Remove(10, 1);
                                                Data2 = Data2.Insert(10, "1");
                                            }
                                            else
                                            {
                                                Data2 = Data2.Remove(10, 1);
                                                Data2 = Data2.Insert(10, "0");
                                            }
                                            break;
                                        case "BUZZER2":

                                            if (val.ToUpper().Equals("TRUE"))
                                            {
                                                Data2 = Data2.Remove(11, 1);
                                                Data2 = Data2.Insert(11, "1");
                                            }
                                            else
                                            {
                                                Data2 = Data2.Remove(11, 1);
                                                Data2 = Data2.Insert(11, "0");
                                            }
                                            break;
                                    }
                                }
                                Data1 = "11111111111000000000000000000000";
                                SIGSTAT_SYSTEM_Data1 = BinaryStringToHexString(Data1);
                                SIGSTAT_SYSTEM_Data2 = BinaryStringToHexString(Data2);
                                if (Events.SYSTEM)
                                {
                                    SendEvent(EventHandling, "SIGSTAT", "SYSTEM", SIGSTAT_SYSTEM_Data1, SIGSTAT_SYSTEM_Data2);
                                }
                                break;
                            case "PORT":
                                Data1 = "00000000000000000000000000000000";
                                Data2 = "00000000000000000000000000000000";
                                Node port = NodeManagement.Get(Name);
                                if (port.Foup_Placement)
                                {
                                    Data1 = Data1.Remove(0, 1);
                                    Data1 = Data1.Insert(0, "1");

                                    Data2 = Data2.Remove(1, 1);
                                    Data2 = Data2.Insert(1, "1");
                                }
                                else
                                {
                                    Data1 = Data1.Remove(0, 1);
                                    Data1 = Data1.Insert(0, "0");

                                    Data2 = Data2.Remove(1, 1);
                                    Data2 = Data2.Insert(1, "0");
                                }

                                if (port.Foup_Presence)
                                {
                                    Data1 = Data1.Remove(1, 1);
                                    Data1 = Data1.Insert(1, "1");

                                    Data2 = Data2.Remove(0, 1);
                                    Data2 = Data2.Insert(0, "1");
                                }
                                else
                                {
                                    Data1 = Data1.Remove(1, 1);
                                    Data1 = Data1.Insert(1, "0");

                                    Data2 = Data2.Remove(0, 1);
                                    Data2 = Data2.Insert(0, "0");
                                }

                                if (port.Access_SW)
                                {
                                    Data1 = Data1.Remove(2, 1);
                                    Data1 = Data1.Insert(2, "1");
                                }
                                else
                                {
                                    Data1 = Data1.Remove(2, 1);
                                    Data1 = Data1.Insert(2, "0");
                                }

                                if (port.Foup_Lock)
                                {
                                    Data1 = Data1.Remove(3, 1);
                                    Data1 = Data1.Insert(3, "1");
                                }
                                else
                                {
                                    Data1 = Data1.Remove(3, 1);
                                    Data1 = Data1.Insert(3, "0");
                                }

                                switch (port.Load_LED)
                                {
                                    case "TRUE":
                                    case "BLINK":
                                        Data2 = Data2.Remove(2, 1);
                                        Data2 = Data2.Insert(2, "1");
                                        break;
                                    case "FALSE":
                                        Data2 = Data2.Remove(2, 1);
                                        Data2 = Data2.Insert(2, "0");
                                        break;
                                }

                                switch (port.UnLoad_LED)
                                {
                                    case "TRUE":
                                    case "BLINK":
                                        Data2 = Data2.Remove(3, 1);
                                        Data2 = Data2.Insert(3, "1");
                                        break;
                                    case "FALSE":
                                        Data2 = Data2.Remove(3, 1);
                                        Data2 = Data2.Insert(3, "0");
                                        break;
                                }

                                switch (port.AccessSW_LED)
                                {
                                    case "TRUE":
                                    case "BLINK":
                                        Data2 = Data2.Remove(8, 1);
                                        Data2 = Data2.Insert(8, "1");
                                        break;
                                    case "FALSE":
                                        Data2 = Data2.Remove(8, 1);
                                        Data2 = Data2.Insert(8, "0");
                                        break;
                                }
                                int pNo = 0;
                                int.TryParse(port.Name.Replace("LOADPORT", ""), out pNo);


                                SIGSTAT_PORT_Data1 = BinaryStringToHexString(Data1);
                                SIGSTAT_PORT_Data2 = BinaryStringToHexString(Data2);
                                if (Events.PORT)
                                {
                                    SendEvent(EventHandling, "SIGSTAT", "P" + pNo.ToString(), SIGSTAT_PORT_Data1, SIGSTAT_PORT_Data2);
                                }
                                break;
                        }
                        break;
                }
            }
            catch (Exception e)
            {
                logger.Error(e.StackTrace);
            }
        }

        public string BinaryStringToHexString(string binary)
        {

            string result = "EEEEEEEE";
            try
            {
                result = BToHex(BinaryToByte(ReverseBit(binary)));
            }
            catch (Exception e)
            {
                logger.Error("BinaryStringToHexString fail: " + binary + "\n" + e.StackTrace);
            }
            return result;

        }

        public string ReverseBit(string BinaryString)
        {
            string result = "";
            for (int i = 0; i < BinaryString.Length; i++)
            {
                result = BinaryString[i] + result;
            }

            return result;
        }

        public byte[] BinaryToByte(string BinaryString)
        {
            byte[] byteOut = new byte[BinaryString.Length / 8];
            for (int i = 0; i < BinaryString.Length; i = i + 8)
            {
                byteOut[i / 8] = Convert.ToByte(BinaryString.Substring(i, 8), 2);
            }
            return byteOut;
        }

        public string BToHex(byte[] BData)
        {
            return BitConverter.ToString(BData).Replace("-", "");
        }

        //***************Event report from Transfer control*****************End
        #endregion
    }
}
