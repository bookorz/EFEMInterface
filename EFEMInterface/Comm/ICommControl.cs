using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFEMInterface.Comm
{
    public interface ICommControl
    {
        void Send(string msg);
    }
}
