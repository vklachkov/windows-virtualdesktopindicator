using VirtualDesktopIndicator.Native;

namespace VirtualDesktopIndicator.Components;

public class DesktopNameForm : Form
{
    protected override CreateParams CreateParams
    {
        get
        {
            var pParams = base.CreateParams;
            pParams.ExStyle |= (int) WindowStylesEx.WS_EX_TOOLWINDOW;
            return pParams;
        }
    }

    private readonly Label _displayLabel;
    
    public DesktopNameForm(string font)
    {
        ShowInTaskbar = false;
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        TopMost = true;
        Height = 100;
        
        Controls.Add(_displayLabel = new()
        {
            Font = new(font, 48, FontStyle.Bold, GraphicsUnit.Pixel),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter
        });
    }

    public void Show(string displayName, Color foreColor, Color backColor)
    {
        _displayLabel.Text = displayName;
        _displayLabel.ForeColor = foreColor;
        
        Opacity = 1;
        BackColor = backColor;

        Width = TextRenderer.MeasureText(_displayLabel.Text, _displayLabel.Font).Width;
        Location = new(Screen.PrimaryScreen.Bounds.Width - Width,
            Screen.PrimaryScreen.Bounds.Height - Height);
        
        Show();
    }
}