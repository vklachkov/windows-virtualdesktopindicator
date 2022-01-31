using System.Reflection;

namespace VirtualDesktopIndicator;

public static class Constants
{
    public static string AppName => Assembly.GetExecutingAssembly().GetName().Name ?? "VirtualDesktopIndicator";
}