﻿using System.Drawing.Drawing2D;
using System.Drawing.Text;
using Microsoft.Win32;
using VirtualDesktopIndicator.Helpers;
using VirtualDesktopIndicator.Native;
using VirtualDesktopIndicator.Native.Constants;
using VirtualDesktopIndicator.Native.Hooks;
using VirtualDesktopIndicator.Native.VirtualDesktop;
using VirtualDesktopIndicator.Utils;
using Timer = System.Windows.Forms.Timer;

namespace VirtualDesktopIndicator.Components;

internal class DesktopNotifyIcon : IDisposable
{
    #region Theme

    private const string RegistryThemeDataPath =
        @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

    private enum Theme
    {
        Light,
        Dark
    }

    private static readonly Dictionary<Theme, Color> ThemesColors = new()
    {
        {Theme.Dark, Color.White},
        {Theme.Light, Color.Black},
    };

    private Color CurrentThemeColor => ThemesColors[_systemTheme];

    private Theme _cachedSystemTheme;
    private Theme _systemTheme;

    private RegistryMonitor? _registryMonitor;

    #endregion

    #region Drawing Constants

    private const string FontName = "Tahoma";

    private static int BorderThinkness => Width / BaseWidth;
    private static int FontSize => (int) Math.Ceiling(Width / 1.5);

    private readonly FontStyle FontStyle = FontStyle.Regular;

    // Default windows tray icon size
    private const int BaseHeight = 16;
    private const int BaseWidth = 16;

    // We use half the size, because otherwise the image is rendered with incorrect anti-aliasing
    private static int Height
    {
        get
        {
            var height = User32.GetSystemMetrics(SystemMetric.SM_CYICON) / 2;
            return height < BaseHeight ? BaseHeight : height;
        }
    }

    private static int Width
    {
        get
        {
            var width = User32.GetSystemMetrics(SystemMetric.SM_CXICON) / 2;
            return width < BaseWidth ? BaseWidth : width;
        }
    }

    #endregion

    private readonly IVirtualDesktopManager _virtualDesktopManager;

    private readonly MouseScrollHook _mouseHook;

    private readonly NotifyIcon _notifyIcon;
    private readonly Timer _timer;

    private uint CurrentVirtualDesktop => _virtualDesktopManager.Current() + 1;

    private uint _lastVirtualDesktop;
    private string? _cachedDisplayText;

    public DesktopNotifyIcon(IVirtualDesktopManager virtualDesktop)
    {
        _virtualDesktopManager = virtualDesktop;

        _notifyIcon = new() {ContextMenuStrip = CreateContextMenu()};
        _notifyIcon.Click += OnNotifyIconClick;

        _timer = new() {Enabled = false};
        _timer.Tick += OnTimerTick;

        _mouseHook = new();
        _cachedSystemTheme = _systemTheme = GetSystemTheme();

        InitMouseHook();
        InitRegistryMonitor();
    }

    #region Lifecycles

    #region NotifyIcon

    public void Show()
    {
        _registryMonitor?.Start();

        _notifyIcon.Visible = true;
        _timer.Enabled = true;
    }

    public void Dispose()
    {
        StopRegistryMonitor();
        StopMouseHook();

        _notifyIcon.Dispose();
        _timer.Dispose();
    }

    #endregion

    #region MouseHook

    private void InitMouseHook()
    {
        _mouseHook.Register();
        _mouseHook.MouseScroll += OnMouseScroll;
    }

    private void StopMouseHook()
    {
        _mouseHook.MouseScroll -= OnMouseScroll;
        _mouseHook.Unregister();
    }

    #endregion

    #region RegistryMonitor

    private void InitRegistryMonitor()
    {
        _registryMonitor = new(RegistryThemeDataPath);

        _registryMonitor.RegChanged += OnRegistryChanged;
        _registryMonitor.Error += OnRegistryError;
    }

    private void StopRegistryMonitor()
    {
        if (_registryMonitor == null) return;

        _registryMonitor.Stop();

        _registryMonitor.RegChanged -= OnRegistryChanged;
        _registryMonitor.Error -= OnRegistryError;

        _registryMonitor = null;
    }

    #endregion

    #endregion

    #region Events

    private void OnMouseScroll(object? sender, MouseScrollEventArgs e)
    {
        // We need to call the virtual desktop functions from a STA thread, because of COM quirks
        var thread = new Thread(() =>
        {
            var rect = Shell32.GetNotifyIconRect(_notifyIcon);
            var cursor = Cursor.Position;
            
            if (rect.left <= cursor.X && rect.right >= cursor.X && rect.top <= cursor.Y && rect.bottom >= cursor.Y)
            {
                if (e.Delta < 0)
                {
                    _virtualDesktopManager.SwitchForward();
                }
                else
                {
                    _virtualDesktopManager.SwitchBackward();
                }
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    private void OnTimerTick(object? sender, EventArgs e)
    {
        try
        {
            if (CurrentVirtualDesktop == _lastVirtualDesktop) return;

            _cachedDisplayText = CurrentVirtualDesktop < 100 ? CurrentVirtualDesktop.ToString() : "++";
            _lastVirtualDesktop = CurrentVirtualDesktop;

            RedrawIcon();
            ShowDesktopNameToast();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"{Constants.AppName} encountered an unhandled error:\n{ex}",
                Constants.AppName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );

            Application.Exit();
        }
    }

    private void OnRegistryChanged(object? sender, EventArgs e)
    {
        _systemTheme = GetSystemTheme();
        if (_systemTheme == _cachedSystemTheme) return;

        RedrawIcon();
        _cachedSystemTheme = _systemTheme;
    }

    private void OnRegistryError(object sender, ErrorEventArgs e)
    {
        StopRegistryMonitor();
    }

    private void OnNotifyIconClick(object? sender, EventArgs eventArgs)
    {
        Shell32.ShellExecuteClsid(ClsIds.TaskView);
    }

    #endregion

    #region Icon

    private void RedrawIcon()
    {
        _notifyIcon.Icon = GenerateIcon();
    }

    private Icon? GenerateIcon()
    {
        var font = new Font(FontName, FontSize, FontStyle, GraphicsUnit.Pixel);
        var brush = new SolidBrush(CurrentThemeColor);
        var bitmap = new Bitmap(Width, Height);

        var graphics = Graphics.FromImage(bitmap);

        graphics.SmoothingMode = SmoothingMode.HighSpeed;
        graphics.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;

        graphics.Clear(Color.Transparent);

        // Draw border
        // The g.DrawRectangle always uses anti-aliasing and border looks very poor at such small resolutions
        // Implement own hack!
        var pen = new Pen(CurrentThemeColor, 1);
        for (var thickness = 0; thickness < BorderThinkness; thickness++)
        {
            // Top
            graphics.DrawLine(pen, 0, thickness, Width - 1, thickness);
            // Right
            graphics.DrawLine(pen, thickness, 0, thickness, Height - 1);
            // Left
            graphics.DrawLine(pen, Width - 1 - thickness, 0, Width - 1 - thickness, Height - 1);
            // Bottom
            graphics.DrawLine(pen, 0, Height - 1 - thickness, Width - 1, Height - 1 - thickness);
        }

        // Draw text
        var textSize = graphics.MeasureString(_cachedDisplayText, font);

        // Calculate padding to center the text
        // We can't assume that g.DrawString will round the coordinates correctly, so we do it manually
        var offsetX = (float) Math.Ceiling((Width - textSize.Width) / 2);
        var offsetY = (float) Math.Ceiling((Height - textSize.Height) / 2);

        graphics.DrawString(_cachedDisplayText, font, brush, offsetX, offsetY);

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

    #endregion

    #region Dynamic Controls

    private ContextMenuStrip CreateContextMenu()
    {
        var menu = new ContextMenuStrip();

        var autorunItem = new ToolStripMenuItem("Start with Windows")
        {
            Checked = Autorun.IsActive()
        };
        autorunItem.Click += (_, _) =>
        {
            autorunItem.Checked = !autorunItem.Checked;

            if (autorunItem.Checked)
            {
                Autorun.Enable();
            }
            else
            {
                Autorun.Disable();
            }
        };

        var exitItem = new ToolStripMenuItem("Quit");
        exitItem.Click += (_, _) => Application.Exit();

        menu.Items.Add(autorunItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(exitItem);

        return menu;
    }

    private void ShowDesktopNameToast(int stayTime = 800, int fadeTime = 25, double fadeStep = 0.025)
    {
        var desktopDisplayLabel = new Label
        {
            Text = _virtualDesktopManager.CurrentDisplayName(),
            Font = new(FontName, 48, FontStyle.Bold, GraphicsUnit.Pixel),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter
        };

        var desktopDisplay = new Form();
        desktopDisplay.ShowInTaskbar = false;
        desktopDisplay.BackColor = CurrentThemeColor;
        desktopDisplay.FormBorderStyle = FormBorderStyle.None;
        desktopDisplay.StartPosition = FormStartPosition.Manual;
        desktopDisplay.Controls.Add(desktopDisplayLabel);
        desktopDisplay.TopMost = true;
        desktopDisplay.Opacity = 1;
        desktopDisplay.Width = TextRenderer.MeasureText(desktopDisplayLabel.Text, desktopDisplayLabel.Font).Width;
        desktopDisplay.Height = 100;
        desktopDisplay.Show();

        // Setting location after showing is somehow more reliable with width and height?
        desktopDisplay.Location = new(Screen.PrimaryScreen.Bounds.Width - desktopDisplay.Width,
            Screen.PrimaryScreen.Bounds.Height - desktopDisplay.Height);

        new Thread(() =>
        {
            // Delay animation
            Thread.Sleep(stayTime);

            // Fade out
            while (desktopDisplay.Opacity > 0)
            {
                desktopDisplay.Invoke(() => desktopDisplay.Opacity -= fadeStep);
                Thread.Sleep(fadeTime);
            }

            // Close form forever
            desktopDisplay.Invoke(() => desktopDisplay.Close());
        }).Start();
    }

    #endregion

    #region Static

    private static Theme GetSystemTheme()
    {
        return (int) (Registry.GetValue(RegistryThemeDataPath, "SystemUsesLightTheme", 0) ?? 0) == 1
            ? Theme.Light
            : Theme.Dark;
    }

    #endregion
}