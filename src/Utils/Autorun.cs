using Microsoft.Win32;

namespace VirtualDesktopIndicator.Utils
{
    internal static class Autorun
    {
        private const string RegistryAutorunPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";

        public static void Enable()
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryAutorunPath, true);
            key?.SetValue(Constants.AppName, $"\"{Application.ExecutablePath}\"");
        }

        public static void Disable()
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryAutorunPath, true);
            key?.DeleteValue(Constants.AppName, false);
        }
        
        public static bool IsActive()
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryAutorunPath, false);
            return (key?.GetValue(Constants.AppName) ?? string.Empty).ToString() == $"\"{Application.ExecutablePath}\"";
        }
    }
}
