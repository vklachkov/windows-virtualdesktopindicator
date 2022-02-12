using System.Reflection;

namespace VirtualDesktopIndicator;

public static class Constants
{
    public static string AppName => Assembly.GetExecutingAssembly().GetName().Name ?? "VirtualDesktopIndicator";
    
    public const string ConfigName = "vdi_config.json";
    public const string TaskView = "Task View";
}