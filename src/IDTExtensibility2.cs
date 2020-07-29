using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace PowerPointExtensibility
{
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnConnection([MarshalAs(26)][In] object application, [In] ext_ConnectMode connectMode, [MarshalAs(26)][In] object addInInst, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        [DispId(2)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnDisconnection([In] ext_DisconnectMode removeMode, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        [DispId(3)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnAddInsUpdate([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        [DispId(4)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnStartupComplete([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        [DispId(5)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnBeginShutdown([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);
    }
}
