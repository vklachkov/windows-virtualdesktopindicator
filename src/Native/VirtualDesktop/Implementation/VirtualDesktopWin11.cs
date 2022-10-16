﻿using System.Runtime.InteropServices;

namespace VirtualDesktopIndicator.Native.VirtualDesktop.Implementation;

internal class VirtualDesktopWin11 : IVirtualDesktopManager
{
    #region API

    public uint Current()
    {
        var currentDesktop = DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop(IntPtr.Zero);
        var currentDesktopIndex = DesktopManager.GetDesktopIndex(currentDesktop);

        return (uint) currentDesktopIndex;
    }

    public void SwitchForward()
    {
        var current = DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop(IntPtr.Zero);

        DesktopManager.VirtualDesktopManagerInternal.GetAdjacentDesktop(current, 4, out var adjacent);
        if (adjacent == null) return;
        DesktopManager.VirtualDesktopManagerInternal.SwitchDesktop(IntPtr.Zero, adjacent);
    }

    public void SwitchBackward()
    {
        var current = DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop(IntPtr.Zero);

        DesktopManager.VirtualDesktopManagerInternal.GetAdjacentDesktop(current, 3, out var adjacent);
        if (adjacent == null) return;
        DesktopManager.VirtualDesktopManagerInternal.SwitchDesktop(IntPtr.Zero, adjacent);
    }

    public string CurrentDisplayName()
    {
        return DesktopNameFromDesktop(DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop(IntPtr.Zero));
    }

    #endregion

    #region Implementation

    private static string DesktopNameFromDesktop(IVirtualDesktop desktop)
    {
        var desktopName = Microsoft.Win32.Registry
            .GetValue(
                $"HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VirtualDesktops\\Desktops\\{{{desktop.GetId()}}}",
                "Name", null)?.ToString();

        if (string.IsNullOrEmpty(desktopName))
        {
            desktopName = "Desktop " + (DesktopManager.GetDesktopIndex(desktop) + 1);
        }

        return desktopName;
    }

    #region COM Guids

    internal static class Guids
    {
        public static readonly Guid CLSID_ImmersiveShell = new Guid("C2F03A33-21F5-47FA-B4BB-156362A2F239");

        public static readonly Guid CLSID_VirtualDesktopManagerInternal =
            new Guid("C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B");
    }

    #endregion

    #region COM API

    [StructLayout(LayoutKind.Sequential)]
    internal struct Size
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    internal enum APPLICATION_VIEW_CLOAK_TYPE : int
    {
        AVCT_NONE = 0,
        AVCT_DEFAULT = 1,
        AVCT_VIRTUAL_DESKTOP = 2
    }

    internal enum APPLICATION_VIEW_COMPATIBILITY_POLICY : int
    {
        AVCP_NONE = 0,
        AVCP_SMALL_SCREEN = 1,
        AVCP_TABLET_SMALL_SCREEN = 2,
        AVCP_VERY_SMALL_SCREEN = 3,
        AVCP_HIGH_SCALE_FACTOR = 4
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIInspectable)]
    [Guid("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
    internal interface IApplicationView
    {
        int SetFocus();
        int SwitchTo();
        int TryInvokeBack(IntPtr /* IAsyncCallback* */ callback);
        int GetThumbnailWindow(out IntPtr hwnd);
        int GetMonitor(out IntPtr /* IImmersiveMonitor */ immersiveMonitor);
        int GetVisibility(out int visibility);
        int SetCloak(APPLICATION_VIEW_CLOAK_TYPE cloakType, int unknown);

        int GetPosition(ref Guid guid /* GUID for IApplicationViewPosition */,
            out IntPtr /* IApplicationViewPosition** */ position);

        int SetPosition(ref IntPtr /* IApplicationViewPosition* */ position);
        int InsertAfterWindow(IntPtr hwnd);
        int GetExtendedFramePosition(out Rect rect);
        int GetAppUserModelId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        int SetAppUserModelId(string id);
        int IsEqualByAppUserModelId(string id, out int result);
        int GetViewState(out uint state);
        int SetViewState(uint state);
        int GetNeediness(out int neediness);
        int GetLastActivationTimestamp(out ulong timestamp);
        int SetLastActivationTimestamp(ulong timestamp);
        int GetVirtualDesktopId(out Guid guid);
        int SetVirtualDesktopId(ref Guid guid);
        int GetShowInSwitchers(out int flag);
        int SetShowInSwitchers(int flag);
        int GetScaleFactor(out int factor);
        int CanReceiveInput(out bool canReceiveInput);
        int GetCompatibilityPolicyType(out APPLICATION_VIEW_COMPATIBILITY_POLICY flags);
        int SetCompatibilityPolicyType(APPLICATION_VIEW_COMPATIBILITY_POLICY flags);
        int GetSizeConstraints(IntPtr /* IImmersiveMonitor* */ monitor, out Size size1, out Size size2);
        int GetSizeConstraintsForDpi(uint uint1, out Size size1, out Size size2);
        int SetSizeConstraintsForDpi(ref uint uint1, ref Size size1, ref Size size2);
        int OnMinSizePreferencesUpdated(IntPtr hwnd);
        int ApplyOperation(IntPtr /* IApplicationViewOperation* */ operation);
        int IsTray(out bool isTray);
        int IsInHighZOrderBand(out bool isInHighZOrderBand);
        int IsSplashScreenPresented(out bool isSplashScreenPresented);
        int Flash();
        int GetRootSwitchableOwner(out IApplicationView rootSwitchableOwner);
        int EnumerateOwnershipTree(out IObjectArray ownershipTree);
        int GetEnterpriseId([MarshalAs(UnmanagedType.LPWStr)] out string enterpriseId);
        int IsMirrored(out bool isMirrored);
        int Unknown1(out int unknown);
        int Unknown2(out int unknown);
        int Unknown3(out int unknown);
        int Unknown4(out int unknown);
        int Unknown5(out int unknown);
        int Unknown6(int unknown);
        int Unknown7();
        int Unknown8(out int unknown);
        int Unknown9(int unknown);
        int Unknown10(int unknownX, int unknownY);
        int Unknown11(int unknown);
        int Unknown12(out Size size1);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("536D3495-B208-4CC9-AE26-DE8111275BF8")]
    internal interface IVirtualDesktop
    {
        bool IsViewVisible(IApplicationView view);
        Guid GetId();
        IntPtr Unknown1();

        [return: MarshalAs(UnmanagedType.HString)]
        string GetName();

        [return: MarshalAs(UnmanagedType.HString)]
        string GetWallpaperPath();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("B2F925B9-5A0F-4D2E-9F4D-2B1507593C10")]
    internal interface IVirtualDesktopManagerInternal
    {
        int GetCount(IntPtr hWnd);
        void MoveViewToDesktop(IApplicationView view, IVirtualDesktop desktop);
        bool CanViewMoveDesktops(IApplicationView view);
        IVirtualDesktop GetCurrentDesktop(IntPtr hWnd);
        void GetDesktops(IntPtr hWnd, out IObjectArray desktops);

        [PreserveSig]
        int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop? desktop);

        void SwitchDesktop(IntPtr hWnd, IVirtualDesktop desktop);
        IVirtualDesktop CreateDesktop(IntPtr hWnd);
        void MoveDesktop(IVirtualDesktop desktop, IntPtr hWnd, int nIndex);
        void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
        IVirtualDesktop FindDesktop(ref Guid desktopid);
        void Unknown1(IVirtualDesktop desktop, out IntPtr unknown1, out IntPtr unknown2);
        void SetName(IVirtualDesktop desktop, [MarshalAs(UnmanagedType.HString)] string name);
        void SetWallpaperPath(IVirtualDesktop desktop, [MarshalAs(UnmanagedType.HString)] string path);
        void SetAllWallpaperPaths([MarshalAs(UnmanagedType.HString)] string path);
        void Unknown2(IApplicationView pView0, IApplicationView pView1);
        int Unknown3();
        void RemoveAll(bool remove);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
    internal interface IObjectArray
    {
        void GetCount(out int count);
        void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object obj);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
    internal interface IServiceProvider10
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object QueryService(ref Guid service, ref Guid riid);
    }

    #endregion

    #region COM wrapper

    internal static class DesktopManager
    {
        internal static IVirtualDesktopManagerInternal VirtualDesktopManagerInternal;

        static DesktopManager()
        {
            var shell = (IServiceProvider10) Activator.CreateInstance(
                Type.GetTypeFromCLSID(Guids.CLSID_ImmersiveShell));
            VirtualDesktopManagerInternal =
                (IVirtualDesktopManagerInternal) shell.QueryService(Guids.CLSID_VirtualDesktopManagerInternal,
                    typeof(IVirtualDesktopManagerInternal).GUID);
        }

        internal static int GetDesktopIndex(IVirtualDesktop desktop)
        {
            var virtualDesktopIDsObj = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VirtualDesktops", "VirtualDesktopIDs", null);
            var virtualDesktopIDsBytes = new byte[0];
            var virtualDesktopIDs = new List<Guid>();

            if (virtualDesktopIDsObj != null)
            {
                virtualDesktopIDsBytes = (byte[])virtualDesktopIDsObj;
            }
            
            for (int i = 0; i < virtualDesktopIDsBytes.Length / 16; i++)
            {
                var idBytes = virtualDesktopIDsBytes.Skip(i*16).Take(16).ToArray();
                var desktopId = new Guid(idBytes);
                virtualDesktopIDs.Add(desktopId);
            }

            return virtualDesktopIDs.IndexOf(desktop.GetId());
        }
    }

    #endregion

    #endregion
}