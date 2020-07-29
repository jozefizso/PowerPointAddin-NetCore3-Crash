using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PowerPointExtensibility
{
    [ComVisible(true)]
    [Guid("21974C45-68E2-46ED-8700-E59D5B58E54F")]
    [ProgId("PowerPointExtensibility.PowerPointAddin")]
    public class PowerPointAddin : IDTExtensibility2
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Trace.WriteLine($"Addin connected to application. Mode: {connectMode}");
            var type = application.GetType();
            var name = type.FullName;
            var isCom = Marshal.IsComObject(application);

            try
            {
                var unknown = Marshal.GetIUnknownForObject(application);
                //var dispatch = Marshal.GetIDispatchForObject(application);
            }
            catch (Exception ex)
            {
                Trace.TraceError($"Marshal failed. {ex}");
            }
        }

        public void OnDisconnection([In] ext_DisconnectMode removeMode, [In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
            Trace.WriteLine($"Addin disconnecting from application. Mode: {removeMode}");
        }

        public void OnAddInsUpdate([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }

        public void OnStartupComplete([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
            Trace.WriteLine($"Addin startup completed.");
        }

        public void OnBeginShutdown([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }
    }
}
