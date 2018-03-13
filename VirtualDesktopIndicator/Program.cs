using System;
using System.Windows.Forms;

namespace VirtualDesktopIndicator
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (TrayIndicator ti = new TrayIndicator())
            {
                ti.Display();
                Application.Run();
            }
        }
    }
}
