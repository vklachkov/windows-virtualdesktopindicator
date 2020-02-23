using Microsoft.Win32;

namespace VirtualDesktopIndicator.Helpers
{
    public class AutorunManager
    {
        public const string RegistryAutorunPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";

        public static bool GetAutorunStatus(string appName, string exePath)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryAutorunPath, false))
            {
                object keyValue = key.GetValue(appName) ?? "";
                return (keyValue.ToString() == $"\"{exePath}\"");
            }
        }

        public static void AddApplicationToAutorun(string appName, string exePath)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryAutorunPath, true))
            {
                key.SetValue(appName, $"\"{exePath}\"");
            }
        }

        public static void RemoveApplicationFromAutorun(string appName)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryAutorunPath, true))
            {
                key.DeleteValue(appName, false);
            }
        }
    }
}
