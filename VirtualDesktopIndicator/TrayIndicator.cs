using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using VirtualDesktopIndicator.Helpers;

namespace VirtualDesktopIndicator
{
    class TrayIndicator : IDisposable
    {
        private static string AppName =>
            Assembly.GetExecutingAssembly().GetName().Name;

        private NotifyIcon trayIcon;
        private Timer timer;

        #region Virtual Desktops

        private static int CurrentVirtualDesktop =>
            VirtualDesktopApi.Desktop.FromDesktop(VirtualDesktopApi.Desktop.Current) + 1;

        private int previewVirtualDesktop;

        private static int VirtualDesktopsCount =>
            VirtualDesktopApi.Desktop.Count;

        #endregion

        #region Theme

        private const string RegistryThemeDataPath =
            @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

        private enum Theme { Light, Dark }

        private static readonly Dictionary<Theme, Color> ThemesColors = new Dictionary<Theme, Color>()
        {
            { Theme.Dark, Color.White },
            { Theme.Light, Color.Black },
        };

        private Color CurrentThemeColor => ThemesColors[systemTheme];

        private Theme cachedSystemTheme;
        private Theme systemTheme;

        private RegistryMonitor registryMonitor;

        #endregion

        #region Drawing data 

        private const string FontName = "Graph 35+ pix";
        private const int FontSize = 32;
        private FontStyle FontStyle = FontStyle.Regular;

        private const int BaseHeight = 16;
        private int Height => SystemMetricsApi.GetSystemMetrics(SystemMetric.SM_CYICON);

        private const int BaseWidth = 16;
        private int Width => SystemMetricsApi.GetSystemMetrics(SystemMetric.SM_CXICON);

        private int BorderThinkness => (int)Math.Ceiling((Height * Width) / (BaseHeight * BaseWidth) / 2.0);

        string cachedDisplayText;

        #endregion

        public TrayIndicator()
        {
            trayIcon = new NotifyIcon { ContextMenuStrip = CreateContextMenu() };
            trayIcon.Click += TrayIconClick;

            timer = new Timer { Enabled = false };
            timer.Tick += TimerTick;

            InitRegistryMonitor();

            cachedSystemTheme = systemTheme = GetSystemTheme();
        }

        #region Events

        private void TimerTick(object sender, EventArgs e)
        {
            try
            {
                if (CurrentVirtualDesktop == previewVirtualDesktop) return;

                cachedDisplayText = CurrentVirtualDesktop < 100 ? CurrentVirtualDesktop.ToString() : "++";
                previewVirtualDesktop = CurrentVirtualDesktop;

                RedrawIcon();
            }
            catch
            {
                MessageBox.Show(
                    "An unhandled error occured!",
                    "VirtualDesktopIndicator",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                Application.Exit();
            }
        }

        private void TrayIconClick(object sender, EventArgs e)
        {
            /*
            MouseEventArgs me = e as MouseEventArgs;

            if (me.Button == MouseButtons.Left)
                ShowTaskView();
            */
        }

        #endregion

        public void Display()
        {
            registryMonitor.Start();

            trayIcon.Visible = true;
            timer.Enabled = true;
        }

        public void Dispose()
        {
            StopRegistryMonitor();

            trayIcon.Dispose();
            timer.Dispose();
        }

        private void InitRegistryMonitor()
        {
            registryMonitor = new RegistryMonitor(RegistryThemeDataPath);

            registryMonitor.RegChanged += new EventHandler(OnRegistryChanged);
            registryMonitor.Error += new ErrorEventHandler(OnRegistryError);
        }

        private void StopRegistryMonitor()
        {
            if (registryMonitor == null) return;

            registryMonitor.Stop();
            registryMonitor.RegChanged -= new EventHandler(OnRegistryChanged);
            registryMonitor.Error -= new System.IO.ErrorEventHandler(OnRegistryError);
            registryMonitor = null;
        }

        private void RedrawIcon()
        {
            trayIcon.Icon = GenerateIcon();
        }

        public static void ShowTaskView()
        {
            /*
             * Unimplemented!
             * I didn't find a efficient way to launch task viewer.
             * Each of them has problems, but here are some solutions:
             *   1. Run "explorer shell:::{3080F90E-D7AD-11D9-BD98-0000947B0257}"
             *   2. Simulating <Win + Tab>
             */
        }

        private ContextMenuStrip CreateContextMenu()
        {
            var menu = new ContextMenuStrip();

            var autorunItem = new ToolStripMenuItem("Start application at Windows startup")
            {
                Checked = AutorunManager.GetAutorunStatus(AppName, Application.ExecutablePath)
            };
            autorunItem.Click += (sender, e) =>
            {
                autorunItem.Checked = !autorunItem.Checked;

                if (autorunItem.Checked)
                {
                    AutorunManager.AddApplicationToAutorun(AppName, Application.ExecutablePath);
                }
                else
                {
                    AutorunManager.RemoveApplicationFromAutorun(AppName);
                }
            };

            menu.Items.Add(new ToolStripSeparator());

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (sender, e) => Application.Exit();

            menu.Items.Add(autorunItem);
            menu.Items.Add(exitItem);

            return menu;
        }

        private Icon GenerateIcon()
        {
            var font = new Font(FontName, FontSize, FontStyle, GraphicsUnit.Pixel);
            var brush = new SolidBrush(CurrentThemeColor);
            var bitmap = new Bitmap(Width, Height);

            var g = Graphics.FromImage(bitmap);

            g.SmoothingMode = SmoothingMode.HighSpeed;
            g.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;

            g.Clear(Color.Transparent);

            // Draw border
            g.DrawRectangle(
                new Pen(CurrentThemeColor, BorderThinkness),
                new Rectangle(1, 1, Width - 2, Height - 2)
            );

            // Draw text
            var textSize = g.MeasureString(cachedDisplayText, font);

            var offsetX = (Width - textSize.Width) / 2 + BorderThinkness / 2;
            var offsetY = (Height - textSize.Height) / 2 + BorderThinkness / 2;

            g.DrawString(cachedDisplayText, font, brush, offsetX, offsetY);

            // Create icon from bitmap and return it
            // bitmapText.GetHicon() can throw exception
            try
            {
                return Icon.FromHandle(bitmap.GetHicon());
            }
            catch
            {
                return null;
            }
        }

        private Theme GetSystemTheme()
        {
            return (int)Registry.GetValue(RegistryThemeDataPath, "SystemUsesLightTheme", 0) == 1 ?
                     Theme.Light :
                     Theme.Dark;
        }

        private void OnRegistryChanged(object sender, EventArgs e)
        {
            systemTheme = GetSystemTheme();
            if (systemTheme == cachedSystemTheme) return;

            RedrawIcon();
            cachedSystemTheme = systemTheme;
        }

        private void OnRegistryError(object sender, ErrorEventArgs e) => StopRegistryMonitor();
    }
}
