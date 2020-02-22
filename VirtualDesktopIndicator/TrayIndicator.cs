using Microsoft.Win32;
using RegistryUtils;
using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VirtualDesktopIndicator
{
    enum Theme { Light, Dark }

    class TrayIndicator : IDisposable
    {
        

        #region Data

        private bool cachedSystemUsedLightTheme;

        private RegistryMonitor registryMonitor;

        Color DarkThemeColor { get; } = Color.White;
        Color LightThemeColor { get; } = Color.Black;

        Theme Theme = Theme.Dark;

        Color iconColor => (Theme == Theme.Dark) ? DarkThemeColor : LightThemeColor;


        string appName = "VirtualDesktopIndicator";

        NotifyIcon trayIcon;
        Timer timer;

        #region Font

        string IconFontName { get; } = "Pixel FJVerdana";
        int IconFontSize { get; } = 10;
        FontStyle IconFontStyle { get; } = FontStyle.Regular;

        #endregion

        #region Common

        int VirtualDesktopsCount => VirtualDesktop.Desktop.Count;

        int CurrentVirtualDesktop => VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current) + 1;
        int CachedVirtualDesktop = 0;

        Color MainColor { get; } = Color.White;

        float OffsetX { get; } = 1;
        float OffsetY { get; } = 1;

        int MagicSize { get; } = 16;  // Constant tray icon size 

        #endregion

        #endregion

        public TrayIndicator()
        {
            trayIcon = new NotifyIcon
            {
                ContextMenuStrip = CreateContextMenu()
            };
            trayIcon.Click += trayIcon_Click;

            timer = new Timer
            {
                Enabled = false
            };
            timer.Tick += timer_Update;

            registryMonitor = new RegistryMonitor(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            registryMonitor.RegChanged += new EventHandler(OnRegChanged);
            registryMonitor.Error += new System.IO.ErrorEventHandler(OnError);
            registryMonitor.Start();
        }

        #region Events

        private void timer_Update(object sender, EventArgs e)
        {
            try
            {
                if (CurrentVirtualDesktop != CachedVirtualDesktop)
                {
                    string iconText = CurrentVirtualDesktop.ToString("00");
                    if (CurrentVirtualDesktop >= 100) iconText = "++";

                    // GenerateIcon() can return null
                    trayIcon.Icon = GenerateIcon(iconText);

                    CachedVirtualDesktop = CurrentVirtualDesktop;
                }
            }
            catch
            {
                MessageBox.Show("An unhandled error occured.", "VirtualDesktopIndicator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        private void trayIcon_Click(object sender, EventArgs e)
        {
            MouseEventArgs me = e as MouseEventArgs;
            if (me.Button == MouseButtons.Left) ShowTaskView();
        }

        #endregion

        #region Functions

        #region Task View

        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        public static void ShowTaskView()
        {
            const int KEYEVENTF_EXTENDEDKEY = 0x0001;  // Key down flag
            const int KEYEVENTF_KEYUP = 0x0002;  // Key up flag

            const int VK_TAB = 0x09;  // Tab key code
            const int VK_LWIN = 0x5B;  // Windows key code

            // Hold Windows down and press Tab
            keybd_event(VK_LWIN, 0, KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0);
        }

        #endregion

        #region Autorun

        private bool GetAutorunStatus()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", false))
            {
                object keyValue = key.GetValue(appName) ?? "";

                if (keyValue.ToString() != "\"" + Application.ExecutablePath + "\"")
                {
                    return false;
                }
                else return true;
            }
        }

        // https://www.fluxbytes.com/csharp/start-application-at-windows-startup/

        private void AddApplicationToAutorun()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                key.SetValue(appName, "\"" + Application.ExecutablePath + "\"");
            }
        }

        private void RemoveApplicationFromAutorun()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                key.DeleteValue(appName, false);
            }
        }

        #endregion

        public void Display()
        {
            trayIcon.Visible = true;
            timer.Enabled = true;
        }

        private ContextMenuStrip CreateContextMenu()
        {
            var menu = new ContextMenuStrip();

            // Autostartup
            ToolStripMenuItem autostartup = new ToolStripMenuItem("Start application at Windows startup")
            {
                Checked = GetAutorunStatus()
            };
            autostartup.Click += (sender, e) =>
                {
                    autostartup.Checked = !autostartup.Checked;

                    if (GetAutorunStatus())
                    {
                        RemoveApplicationFromAutorun();
                    }
                    else
                    {
                        AddApplicationToAutorun();
                    }
                };
            menu.Items.Add(autostartup);

            // Separator
            menu.Items.Add(new ToolStripSeparator());

            // Exit
            ToolStripMenuItem exit = new ToolStripMenuItem("Exit");
            exit.Click += (sender, e) => Application.Exit();
            menu.Items.Add(exit);

            return menu;
        }

        private Icon GenerateIcon(string text)
        {
            Font fontToUse = new Font(IconFontName, IconFontSize, IconFontStyle, GraphicsUnit.Pixel);
            Brush brushToUse = new SolidBrush(MainColor);
            Bitmap bitmapText = new Bitmap(MagicSize, MagicSize);  // Const size for tray icon
            Brush brushToUse = new SolidBrush(iconColor);

            Graphics g = Graphics.FromImage(bitmapText);

            g.Clear(Color.Transparent);

            // Draw border
            g.DrawRectangle(
                new Pen(MainColor, 1),
                new Rectangle(0, 0, MagicSize - 1, MagicSize - 1));
                new Pen(iconColor, 1),

            // Draw text
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit;
            g.DrawString(text, fontToUse, brushToUse, OffsetX, OffsetY);

            // Create icon from bitmap and return it
            // bitmapText.GetHicon() can throw exception
            try
            {
                return Icon.FromHandle(bitmapText.GetHicon());
            }
            catch
            {
                return null;
            }
        }

        #endregion

        private void OnRegChanged(object sender, EventArgs e)
        {
            var systemUsedLightTheme = (int)Registry.GetValue(
                @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                "SystemUsesLightTheme",
                0
            ) == 1;

            if (systemUsedLightTheme != cachedSystemUsedLightTheme)
            {
                Theme = systemUsedLightTheme ? Theme.Light : Theme.Dark;

            }

            trayIcon.Icon = GenerateIcon(iconText);

            cachedSystemUsedLightTheme = systemUsedLightTheme;
        }

        private void OnError(object sender, ErrorEventArgs e)
        {
            if (registryMonitor != null)
            {
                registryMonitor.Stop();
                registryMonitor.RegChanged -= new EventHandler(OnRegChanged);
                registryMonitor.Error -= new System.IO.ErrorEventHandler(OnError);
                registryMonitor = null;
            }
        }

        public void Dispose()
        {
            trayIcon.Dispose();
            timer.Dispose();
        }
    }
}
