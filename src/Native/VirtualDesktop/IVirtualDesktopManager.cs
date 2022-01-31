namespace VirtualDesktopIndicator.Native.VirtualDesktop
{
    interface IVirtualDesktopManager
    {
        uint Current();

        void SwitchForward();

        void SwitchBackward();

        string CurrentDisplayName();
    }
}
