using Microsoft.Win32;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace VirtualDesktopIndicator
{
    class TrayIndicator : IDisposable
    {
        #region Data

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

            timer = new Timer
            {
                Enabled = false
            };
            timer.Tick += timer_Update;
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
                Application.Restart();
            }
        }

        #endregion

        #region Functions

        #region Autostartup

        public static void AddApplicationToStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                key.SetValue("My Program", "\"" + Application.ExecutablePath + "\"");
            }
        }

        public static void RemoveApplicationFromStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                key.DeleteValue("My Program", false);
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

            Graphics g = Graphics.FromImage(bitmapText);

            g.Clear(Color.Transparent);

            // Draw border
            g.DrawRectangle(
                new Pen(MainColor, 1),
                new Rectangle(0, 0, MagicSize - 1, MagicSize - 1));

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

        public void Dispose()
        {
            trayIcon.Dispose();
            timer.Dispose();
        }
    }
}
