using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace EFEMInterface.MessageInterface
{
    public interface ICommMessage
    {
        void On_Connection_Message(Socket handler,string content);
        void On_Connection_Connecting();
        void On_Connection_Connected();
        void On_Connection_Disconnected();
        void On_Connection_Error(string Msg);
        
    }
}
