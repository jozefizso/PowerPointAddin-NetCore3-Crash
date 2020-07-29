using System;
using System.Runtime.InteropServices;

namespace PowerPointExtensibility
{
    [Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }
}
