using VirtualDesktopIndicator.Native;
using Timer = System.Windows.Forms.Timer;

namespace VirtualDesktopIndicator.Components;

public class DesktopNameForm : Form
{
    
    private const double FADE_STEP = 0.025;
    private const int DELAY_TICKS = 15;
    
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
    private readonly Timer _animationTimer;

    private int _delayTicks;
    
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

        _animationTimer = new()
        {
            Interval = 75
        };
        _animationTimer.Tick += AnimationTimer_OnTick;
    }

    private void AnimationTimer_OnTick(object? sender, EventArgs e)
    {
        if (_delayTicks-- > 0)
        {
            return;
        }
        
        if (Opacity > 0)
        {
            Opacity -= FADE_STEP;
        }
        else
        {
            _animationTimer.Stop();
            Hide();
        }
    }

    public void Show(string displayName, Color foreColor, Color backColor)
    {
        _displayLabel.Text = displayName;
        _displayLabel.ForeColor = foreColor;
        
        Opacity = 0.5;
        BackColor = backColor;

        Width = TextRenderer.MeasureText(_displayLabel.Text, _displayLabel.Font).Width;
        Location = new(Screen.PrimaryScreen.Bounds.Width - Width,
            Screen.PrimaryScreen.Bounds.Height - Height);

        _delayTicks = DELAY_TICKS;
        _animationTimer.Start();
        
        Show();
    }
}