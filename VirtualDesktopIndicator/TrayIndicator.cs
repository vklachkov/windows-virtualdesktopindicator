using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VirtualDesktopIndicator
{
    class TrayIndicator : IDisposable
    {
        #region Data

        NotifyIcon notifyIcon;
        Timer timer;

        #region Font

        string IconFontName { get; } = "Pixel FJVerdana";
        int IconFontSize { get; } = 10;
        FontStyle IconFontStyle { get; } = FontStyle.Regular;

        #endregion

        #region Common

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
            notifyIcon = new NotifyIcon();

            timer = new Timer();

            timer.Enabled = false;
            timer.Tick += timerUpdate;
        }

        public void Display()
        {
            notifyIcon.Visible = true;
            timer.Enabled = true;
        }

        private void timerUpdate(object sender, EventArgs e)
        {
            if (CurrentVirtualDesktop != CachedVirtualDesktop)
            {
                string iconText = CurrentVirtualDesktop.ToString("00");
                if (CurrentVirtualDesktop >= 100) iconText = "++";

                try
                {
                    notifyIcon.Icon = GenerateIcon(iconText);
                }
                catch
                {

                }

                CachedVirtualDesktop = CurrentVirtualDesktop;
            }
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
            return Icon.FromHandle(bitmapText.GetHicon());
        }

        public void Dispose()
        {
            notifyIcon.Dispose();
            timer.Dispose();
        }
    }
}
