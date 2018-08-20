using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace EFEMInterface
{
    public interface IEFEMControl
    {
        void On_Connection_Connecting();
        void On_Connection_Connected();
        void On_Connection_Disconnected();
        void On_CommandMessage(string msg);

    }
}
