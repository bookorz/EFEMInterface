using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using static EFEMInterface.MessageInterface.RorzeInterface;

namespace EFEMInterface.MessageInterface
{
    public class OnHandling
    {
        public string ID { get; set; }
        public RorzeCommand Cmd { get; set; }
        public Socket Handler { get; set; }
        public int INF_RetryCount { get; set; }
        public DateTime ReceiveTime { get; set; }
        //逾時
        private System.Timers.Timer timeOutTimer = new System.Timers.Timer();
        IHandlingTimeOutReport _TimeOutReport;
        public OnHandling(IHandlingTimeOutReport TimeOutReport,int TimeOut = 5000)
        {
            _TimeOutReport = TimeOutReport;
            timeOutTimer.Enabled = false;
            timeOutTimer.Interval = TimeOut;
            timeOutTimer.Elapsed += new System.Timers.ElapsedEventHandler(TimeOutMonitor);
            INF_RetryCount = 0;
            ID = Guid.NewGuid().ToString();
            ReceiveTime = DateTime.Now;
        }

        public void SetTimeOutMonitor(bool Enabled)
        {

            if (Enabled)
            {
                timeOutTimer.Start();
            }
            else
            {
                timeOutTimer.Stop();
            }

        }

        private void TimeOutMonitor(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (_TimeOutReport != null)
            {
                _TimeOutReport.On_Handling_TimeOut(this);
            }
        }
    }
}
