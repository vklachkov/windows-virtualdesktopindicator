using VirtualDesktopIndicator.Components;
using VirtualDesktopIndicator.Native.VirtualDesktop;
using VirtualDesktopIndicator.Native.VirtualDesktop.Implementation;

namespace VirtualDesktopIndicator;

public static class Program
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        using var notifyIcon = new DesktopNotifyIcon(GetVirtualDesktop());
        notifyIcon.Show();
            
        Application.Run();
    }

    private static IVirtualDesktopManager GetVirtualDesktop()
    {
        const int win11MinBuild = 22000;
        return Environment.OSVersion.Version.Build >= win11MinBuild ? new VirtualDesktopWin11() : new VirtualDesktopWin10();
    }
}