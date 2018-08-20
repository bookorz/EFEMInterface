using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EFEMInterface.MessageInterface.RorzeInterface;

namespace EFEMInterface.MessageInterface
{
    public interface IHandlingTimeOutReport
    {
        void On_Handling_TimeOut(OnHandling TimeOutCmd);
    }
}
