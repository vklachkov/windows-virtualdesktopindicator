using System;
using System.Windows.Forms;
using VirtualDesktopIndicator.Api;

namespace VirtualDesktopIndicator
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (TrayIndicator ti = new TrayIndicator(GetActualDesktopApi()))
            {
                ti.Display();
                Application.Run();
            }
        }

        private static IVirtualDesktopApi GetActualDesktopApi()
        {
            return new Latest();
        }
    }
}
