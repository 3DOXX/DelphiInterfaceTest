using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace openxml2
{
    [ComVisible(true)]
    [Guid("07056E7A-183B-49B5-8391-F32DBC68ECC4"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IExcelLib2
    {
        void Test(out uint result);
    }
}
