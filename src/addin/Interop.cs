using System;
using System.Runtime.InteropServices;

namespace Extensibility
{
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [TypeLibType(TypeLibTypeFlags.FDual | TypeLibTypeFlags.FDispatchable)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            ext_ConnectMode ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            ref Array custom);

        [DispId(2)]
        void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);

        [DispId(4)]
        void OnStartupComplete(ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
    }
}

namespace Microsoft.Office.Core
{
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [TypeLibType(TypeLibTypeFlags.FDual | TypeLibTypeFlags.FDispatchable)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI(string RibbonID);
    }

    [ComImport]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonControl
    {
        [DispId(1)]
        string Id { get; }
        [DispId(2)]
        object Context { get; }
        [DispId(3)]
        string Tag { get; }
    }

    [ComImport]
    [Guid("000C0397-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonUI
    {
        [DispId(1)]
        void Invalidate();
        [DispId(2)]
        void InvalidateControl(string controlID);
    }
}
