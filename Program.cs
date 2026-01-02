// Program.cs
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers; // BASIC
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics;      // ★追加：ヘルプリンクでブラウザ起動
using System.Linq;             // ★追加：Attributes 取得用

internal static class Program
{
    [STAThread]
    static void Main()
    {
        I18n.SetLanguage(SettingsStore.Current.Language);

        ApplicationConfiguration.Initialize();
        Application.Run(new TrayAppContext());
    }
}

internal sealed class AppSettings
{
    public string BaseUrl { get; set; } = "";
    public string TokenEncrypted { get; set; } = "";

    public bool Enabled { get; set; } = true;
    public bool ShowMessageOnSuccess { get; set; } = true;

    public uint HotkeyModifiers { get; set; } = 0x0001 | 0x0002; // Alt|Ctrl
    public int HotkeyVk { get; set; } = 0x4E; // N
    public string HotkeyDisplay { get; set; } = "Ctrl + Alt + N";

    public string CleanupBaseUrl { get; set; } = "";
    public string CleanupTokenEncrypted { get; set; } = "";
    public bool CleanupPretty { get; set; } = true;

    public bool CleanupDailyEnabled { get; set; } = false;
    public int CleanupDailyHour { get; set; } = 2;
    public int CleanupDailyMinute { get; set; } = 0;

    public bool CleanupEveryEnabled { get; set; } = false;
    public int CleanupEveryMinutes { get; set; } = 60;

    public string Language { get; set; } = "en"; // "en" / "ja" / "tr"

    // BASIC
    public string BasicUser { get; set; } = "";
    public string BasicPassEncrypted { get; set; } = "";
}

internal static class SettingsStore
{
    private static readonly object _lock = new();

    private static readonly string DirPath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ClipboardUrlSender");

    private static readonly string FilePath = Path.Combine(DirPath, "settings.json");

    public static AppSettings Current { get; private set; } = Load();

    public static event EventHandler<AppSettings>? Saved;

    public static AppSettings Load()
    {
        lock (_lock)
        {
            try
            {
                if (!File.Exists(FilePath))
                    return new AppSettings();

                var json = File.ReadAllText(FilePath);
                return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }
    }

    public static void Save(AppSettings settings)
    {
        lock (_lock)
        {
            Directory.CreateDirectory(DirPath);
            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(FilePath, json);
            Current = settings;
        }

        Saved?.Invoke(null, settings);
    }
}

internal static class EmbeddedIcon
{
    public static System.Drawing.Icon LoadByFileName(string fileName)
    {
        var asm = Assembly.GetExecutingAssembly();
        foreach (var name in asm.GetManifestResourceNames())
        {
            if (name.EndsWith("." + fileName, StringComparison.OrdinalIgnoreCase) ||
                name.EndsWith("/" + fileName, StringComparison.OrdinalIgnoreCase) ||
                name.EndsWith("\\" + fileName, StringComparison.OrdinalIgnoreCase) ||
                name.EndsWith(fileName, StringComparison.OrdinalIgnoreCase))
            {
                var stream = asm.GetManifestResourceStream(name);
                if (stream == null) continue;

                using (stream)
                {
                    return new System.Drawing.Icon(stream);
                }
            }
        }

        throw new FileNotFoundException(
            $"Embedded resource not found: {fileName}\n" +
            $"確認: ico のビルドアクションが『埋め込みリソース』になっているか");
    }
}

// BASIC
internal static class BasicAuth
{
    public static void Apply(HttpRequestMessage req, AppSettings s)
    {
        var user = (s.BasicUser ?? "").Trim();
        if (string.IsNullOrEmpty(user)) return;

        var pass = DpapiHelper.Decrypt(s.BasicPassEncrypted) ?? "";
        var raw = $"{user}:{pass}";
        var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(raw));
        req.Headers.Authorization = new AuthenticationHeaderValue("Basic", b64);
    }
}

internal sealed class TrayAppContext : ApplicationContext
{
    // ★ここをあなたの GitHub / Docs URL に変更
    private const string HELP_URL = "https://github.com/yourname/ClipboardSender";

    private readonly NotifyIcon _tray;
    private readonly ClipboardWatcherForm _watcher;
    private readonly ToolStripMenuItem _enabledMenuItem;

    private readonly System.Drawing.Icon _iconOn;
    private readonly System.Drawing.Icon _iconOff;

    private readonly ToolStripMenuItem _cleanupModeInfoItem;
    private readonly ToolStripMenuItem _cleanupLastResultItem;
    private DateTime? _cleanupLastAt;
    private bool? _cleanupLastOk;

    private readonly ToolStripMenuItem _menuDeleteInbox;
    private readonly ToolStripMenuItem _menuSettings;

    // ★追加：Help / About
    private readonly ToolStripMenuItem _menuHelp;
    private readonly ToolStripMenuItem _menuAbout;

    private readonly ToolStripMenuItem _menuExit;

    public TrayAppContext()
    {
        _watcher = new ClipboardWatcherForm();

        _iconOn = EmbeddedIcon.LoadByFileName("tray_on.ico");
        _iconOff = EmbeddedIcon.LoadByFileName("tray_off.ico");

        _tray = new NotifyIcon
        {
            Icon = SettingsStore.Current.Enabled ? _iconOn : _iconOff,
            Text = I18n.T("TrayTitle"),
            Visible = true,
            ContextMenuStrip = new ContextMenuStrip()
        };

        _tray.MouseClick += (_, e) =>
        {
            if (e.Button == MouseButtons.Left) ToggleEnabledFromTray();
        };

        _enabledMenuItem = new ToolStripMenuItem { CheckOnClick = false };
        _enabledMenuItem.Click += (_, __) => ToggleEnabledFromTray();

        _cleanupModeInfoItem = new ToolStripMenuItem { Enabled = false };
        _cleanupLastResultItem = new ToolStripMenuItem { Enabled = false };

        _menuDeleteInbox = new ToolStripMenuItem();
        _menuSettings = new ToolStripMenuItem();
        _menuHelp = new ToolStripMenuItem();   // ★追加
        _menuAbout = new ToolStripMenuItem();  // ★追加
        _menuExit = new ToolStripMenuItem();

        _tray.ContextMenuStrip.Items.Add(_enabledMenuItem);
        _tray.ContextMenuStrip.Items.Add(new ToolStripSeparator());

        _tray.ContextMenuStrip.Items.Add(_cleanupModeInfoItem);
        _tray.ContextMenuStrip.Items.Add(_cleanupLastResultItem);
        _tray.ContextMenuStrip.Items.Add(new ToolStripSeparator());

        _tray.ContextMenuStrip.Items.Add(_menuDeleteInbox);
        _tray.ContextMenuStrip.Items.Add(_menuSettings);

        // ★追加：Help / About を Settings の下に
        _tray.ContextMenuStrip.Items.Add(_menuHelp);
        _tray.ContextMenuStrip.Items.Add(_menuAbout);

        _tray.ContextMenuStrip.Items.Add(new ToolStripSeparator());
        _tray.ContextMenuStrip.Items.Add(_menuExit);

        _menuDeleteInbox.Click += async (_, __) =>
        {
            var r = MessageBox.Show(
                I18n.T("ConfirmDeleteInboxBody"),
                I18n.T("ConfirmTitle"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            if (r != DialogResult.Yes) return;

            _tray.BalloonTipTitle = I18n.T("CleanupTitle");
            _tray.BalloonTipText = I18n.T("CleanupRunning");
            _tray.ShowBalloonTip(1000);

            var (ok, info) = await CleanupApi.DeleteInboxAllAsync();

            using var top = new Form { TopMost = true, ShowInTaskbar = false };
            top.StartPosition = FormStartPosition.Manual;
            top.Location = new System.Drawing.Point(-2000, -2000);
            top.Show();

            MessageBox.Show(top, info,
                ok ? I18n.T("CleanupDoneTitle") : I18n.T("CleanupFailTitle"),
                MessageBoxButtons.OK,
                ok ? MessageBoxIcon.Information : MessageBoxIcon.Error);
        };

        _menuSettings.Click += (_, __) =>
        {
            using var f = new SettingsForm();
            f.ShowDialog();

            SyncEnabledMenu();
            _watcher.ApplyHotkeyFromSettings();
            CleanupScheduler.ApplyFromSettings();
            UpdateCleanupModeInfo();
        };

        // ★追加：Help（ブラウザ起動）
        _menuHelp.Click += (_, __) =>
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = HELP_URL,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, I18n.T("MenuHelp"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        };

        // ★追加：About（バージョン表示）
        _menuAbout.Click += (_, __) =>
        {
            using var top = new Form { TopMost = true, ShowInTaskbar = false };
            top.StartPosition = FormStartPosition.Manual;
            top.Location = new System.Drawing.Point(-2000, -2000);
            top.Show();

            var appName = AppInfo.GetProductName();
            var ver = AppInfo.GetVersionString();
            var body = string.Format(I18n.T("AboutBodyFormat"), appName, ver, HELP_URL);

            MessageBox.Show(top, body, I18n.T("AboutTitle"),
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        };

        _menuExit.Click += (_, __) => ExitThread();

        _watcher.EnabledToggled += (_, enabled) =>
        {
            SyncEnabledMenu();
            ShowToggleBalloon(enabled);
        };

        _watcher.ClipboardTextCopied += async (_, text) =>
        {
            if (!SettingsStore.Current.Enabled) return;

            var (ok, info) = await Sender.SendAsync(text);
            var showSuccess = SettingsStore.Current.ShowMessageOnSuccess;

            if (!ok || showSuccess)
            {
                using var top = new Form { TopMost = true, ShowInTaskbar = false };
                top.StartPosition = FormStartPosition.Manual;
                top.Location = new System.Drawing.Point(-2000, -2000);
                top.Show();

                MessageBox.Show(top, info,
                    ok ? I18n.T("SendOkTitle") : I18n.T("SendFailTitle"),
                    MessageBoxButtons.OK,
                    ok ? MessageBoxIcon.Information : MessageBoxIcon.Error);
            }
        };

        _watcher.Show();
        _watcher.Hide();

        CleanupScheduler.ApplyFromSettings();

        RefreshTrayTexts();
        UpdateCleanupModeInfo();
        UpdateCleanupLastResultInfo();
        SyncEnabledMenu();

        CleanupScheduler.CleanupFinished += (_, result) =>
        {
            var (ok, info) = result;

            _cleanupLastAt = DateTime.Now;
            _cleanupLastOk = ok;

            RunOnUi(() =>
            {
                UpdateCleanupLastResultInfo();

                _tray.BalloonTipTitle = I18n.T("AutoCleanupTitle");
                _tray.BalloonTipText = ok ? I18n.T("AutoCleanupOk") : (I18n.T("AutoCleanupFail") + "\n" + Shorten(info, 120));
                _tray.ShowBalloonTip(1000);
            });
        };

        SettingsStore.Saved += (_, newSettings) =>
        {
            RunOnUi(() =>
            {
                I18n.SetLanguage(newSettings.Language);
                RefreshTrayTexts();
                SyncEnabledMenu();
                UpdateCleanupModeInfo();
            });
        };
    }

    private void RunOnUi(Action a)
    {
        try
        {
            if (_watcher.IsHandleCreated && _watcher.InvokeRequired) _watcher.BeginInvoke(a);
            else a();
        }
        catch { }
    }

    private void RefreshTrayTexts()
    {
        _tray.Text = I18n.T("TrayTitle");

        _menuDeleteInbox.Text = I18n.T("MenuDeleteInbox");
        _menuSettings.Text = I18n.T("MenuSettings");

        // ★追加
        _menuHelp.Text = I18n.T("MenuHelp");
        _menuAbout.Text = I18n.T("MenuAbout");

        _menuExit.Text = I18n.T("MenuExit");

        UpdateCleanupModeInfo();
        UpdateCleanupLastResultInfo();
    }

    private void ToggleEnabledFromTray()
    {
        var s = SettingsStore.Current;
        s.Enabled = !s.Enabled;
        SettingsStore.Save(s);

        SyncEnabledMenu();
        ShowToggleBalloon(s.Enabled);
    }

    private void SyncEnabledMenu()
    {
        _enabledMenuItem.Checked = SettingsStore.Current.Enabled;
        _enabledMenuItem.Text = SettingsStore.Current.Enabled ? I18n.T("MenuEnabled") : I18n.T("MenuDisabled");
        _tray.Icon = SettingsStore.Current.Enabled ? _iconOn : _iconOff;
    }

    private void ShowToggleBalloon(bool enabled)
    {
        _tray.BalloonTipTitle = I18n.T("TrayTitle");
        _tray.BalloonTipText = enabled ? I18n.T("BalloonEnabled") : I18n.T("BalloonDisabled");
        _tray.ShowBalloonTip(1000);
    }

    private void UpdateCleanupModeInfo()
    {
        _cleanupModeInfoItem.Text = string.Format(I18n.T("CleanupModeInfoFormat"), GetCleanupModeText());
    }

    private static string GetCleanupModeText()
    {
        var s = SettingsStore.Current;

        if (s.CleanupDailyEnabled && s.CleanupEveryEnabled)
        {
            return string.Format(I18n.T("CleanupModeDailyAndEveryFormat"),
                s.CleanupDailyHour, s.CleanupDailyMinute, Math.Max(1, s.CleanupEveryMinutes));
        }
        if (s.CleanupDailyEnabled)
        {
            return string.Format(I18n.T("CleanupModeDailyFormat"),
                s.CleanupDailyHour, s.CleanupDailyMinute);
        }
        if (s.CleanupEveryEnabled)
        {
            return string.Format(I18n.T("CleanupModeEveryFormat"),
                Math.Max(1, s.CleanupEveryMinutes));
        }
        return I18n.T("CleanupModeOff");
    }

    private void UpdateCleanupLastResultInfo()
    {
        if (_cleanupLastAt is null || _cleanupLastOk is null)
        {
            _cleanupLastResultItem.Text = I18n.T("CleanupLastNotYet");
            return;
        }

        var at = _cleanupLastAt.Value;
        var ok = _cleanupLastOk.Value;

        _cleanupLastResultItem.Text = string.Format(
            I18n.T("CleanupLastFormat"),
            at.ToString("yyyy-MM-dd HH:mm:ss"),
            ok ? I18n.T("WordSuccess") : I18n.T("WordFail"));
    }

    private static string Shorten(string s, int max)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        s = s.Replace("\r", "").Replace("\n", " ");
        return s.Length <= max ? s : s.Substring(0, max) + "...";
    }

    protected override void ExitThreadCore()
    {
        _tray.Visible = false;
        _tray.Dispose();
        _watcher.Dispose();
        CleanupScheduler.Stop();
        _iconOn.Dispose();
        _iconOff.Dispose();
        base.ExitThreadCore();
    }
}

internal static class AppInfo
{
    public static string GetProductName()
    {
        // AssemblyProduct があればそれ、なければ asm 名
        var asm = Assembly.GetExecutingAssembly();
        var prod = asm.GetCustomAttributes<AssemblyProductAttribute>().FirstOrDefault()?.Product;
        if (!string.IsNullOrWhiteSpace(prod)) return prod!;
        return asm.GetName().Name ?? "ClipboardSender";
    }

    public static string GetVersionString()
    {
        // 1) InformationalVersion (推奨: 1.2.3 / 1.2.3+gitsha)
        var asm = Assembly.GetExecutingAssembly();
        var info = asm.GetCustomAttributes<AssemblyInformationalVersionAttribute>()
                      .FirstOrDefault()?.InformationalVersion;

        if (!string.IsNullOrWhiteSpace(info)) return info!;

        // 2) FileVersion
        var fv = asm.GetCustomAttributes<AssemblyFileVersionAttribute>()
                    .FirstOrDefault()?.Version;
        if (!string.IsNullOrWhiteSpace(fv)) return fv!;

        // 3) AssemblyVersion
        return asm.GetName().Version?.ToString() ?? "0.0.0";
    }
}

internal sealed class ClipboardWatcherForm : Form
{
    private const int WM_HOTKEY = 0x0312;
    private const int HOTKEY_ID_TOGGLE = 1;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, int vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public event EventHandler<string>? ClipboardTextCopied;
    public event EventHandler<bool>? EnabledToggled;

    private string? _lastClipboardText;
    private DateTime _lastEventAtUtc = DateTime.MinValue;

    private static readonly TimeSpan Debounce = TimeSpan.FromMilliseconds(350);
    private bool _hotkeyRegistered = false;

    public void ApplyHotkeyFromSettings()
    {
        if (_hotkeyRegistered)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID_TOGGLE);
            _hotkeyRegistered = false;
        }

        var s = SettingsStore.Current;
        var ok = RegisterHotKey(this.Handle, HOTKEY_ID_TOGGLE, s.HotkeyModifiers, s.HotkeyVk);
        _hotkeyRegistered = ok;
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        ShowInTaskbar = false;
        WindowState = FormWindowState.Minimized;
        Opacity = 0;
        Hide();
    }

    protected override void WndProc(ref Message m)
    {
        const int WM_CLIPBOARDUPDATE = 0x031D;

        if (m.Msg == WM_CLIPBOARDUPDATE)
        {
            _ = HandleClipboardUpdateAsync();
        }
        else if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID_TOGGLE)
        {
            ToggleEnabled();
        }

        base.WndProc(ref m);
    }

    private void ToggleEnabled()
    {
        var s = SettingsStore.Current;
        s.Enabled = !s.Enabled;
        SettingsStore.Save(s);
        EnabledToggled?.Invoke(this, s.Enabled);
    }

    private async Task HandleClipboardUpdateAsync()
    {
        try
        {
            var now = DateTime.UtcNow;
            if (now - _lastEventAtUtc < Debounce) return;

            if (!Clipboard.ContainsText(TextDataFormat.UnicodeText)) return;

            string text = await TryGetClipboardTextAsync();
            text = (text ?? "").Replace("\r\n", "\n");

            if (string.IsNullOrWhiteSpace(text)) return;
            if (text == _lastClipboardText) return;

            _lastClipboardText = text;
            _lastEventAtUtc = now;

            ClipboardTextCopied?.Invoke(this, text);
        }
        catch { }
    }

    private static async Task<string> TryGetClipboardTextAsync()
    {
        for (int i = 0; i < 3; i++)
        {
            try { return Clipboard.GetText(TextDataFormat.UnicodeText); }
            catch { await Task.Delay(40); }
        }
        return "";
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        AddClipboardFormatListener(this.Handle);
        ApplyHotkeyFromSettings();
    }

    protected override void OnHandleDestroyed(EventArgs e)
    {
        UnregisterHotKey(this.Handle, HOTKEY_ID_TOGGLE);
        RemoveClipboardFormatListener(this.Handle);
        base.OnHandleDestroyed(e);
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool AddClipboardFormatListener(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
}

internal static class Sender
{
    private static readonly HttpClient Client = new()
    {
        Timeout = TimeSpan.FromSeconds(20)
    };

    private static string? _lastSent;
    private static DateTime _lastSentAtUtc = DateTime.MinValue;
    private static int _sending = 0;

    public static async Task<(bool ok, string info)> SendAsync(string text)
    {
        if (Interlocked.Exchange(ref _sending, 1) == 1)
            return (false, "busy");

        try
        {
            var s = SettingsStore.Current;

            var baseUrl = (s.BaseUrl ?? "").Trim();
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var _))
                return (false, "settings: invalid url");

            var token = DpapiHelper.Decrypt(s.TokenEncrypted).Trim();
            if (string.IsNullOrWhiteSpace(token))
                return (false, "token is empty or cannot be decrypted");

            text = (text ?? "").Replace("\r\n", "\n");
            if (string.IsNullOrWhiteSpace(text))
                return (false, "empty text");

            if (text == _lastSent && DateTime.UtcNow - _lastSentAtUtc < TimeSpan.FromSeconds(2))
                return (false, "duplicate");

            var kv = new List<KeyValuePair<string, string>>
            {
                new("token", token),
                new("text", text),
            };

            using var content = new FormUrlEncodedContent(kv);
            content.Headers.ContentType!.CharSet = "utf-8";

            using var req = new HttpRequestMessage(HttpMethod.Post, baseUrl) { Content = content };
            BasicAuth.Apply(req, s);

            using var res = await Client.SendAsync(req);
            var body = await res.Content.ReadAsStringAsync();

            _lastSent = text;
            _lastSentAtUtc = DateTime.UtcNow;

            if (!res.IsSuccessStatusCode)
                return (false, $"HTTP {(int)res.StatusCode}\n\n{body}");

            return (true, string.IsNullOrWhiteSpace(body) ? "sent" : body);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
        finally
        {
            Interlocked.Exchange(ref _sending, 0);
        }
    }
}

internal static class CleanupApi
{
    private static readonly HttpClient Client = new()
    {
        Timeout = TimeSpan.FromSeconds(30)
    };

    public static async Task<(bool ok, string info)> DeleteInboxAllAsync()
    {
        try
        {
            var s = SettingsStore.Current;

            var baseUrl = (s.CleanupBaseUrl ?? "").Trim();
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var _))
                return (false, "Cleanup_API URL is not set or invalid");

            var token = DpapiHelper.Decrypt(s.CleanupTokenEncrypted).Trim();
            if (string.IsNullOrWhiteSpace(token))
                return (false, "Cleanup_API token is not set / cannot be decrypted");

            var url = baseUrl;
            if (s.CleanupPretty) url = AppendQuery(baseUrl, "pretty", "1");

            var kv = new List<KeyValuePair<string, string>>
            {
                new("token", token),
                new("confirm", "YES"),
                new("category", "INBOX"),
            };

            using var content = new FormUrlEncodedContent(kv);
            content.Headers.ContentType!.CharSet = "utf-8";

            using var req = new HttpRequestMessage(HttpMethod.Post, url) { Content = content };
            BasicAuth.Apply(req, s);

            using var res = await Client.SendAsync(req);
            var body = await res.Content.ReadAsStringAsync();

            if (!res.IsSuccessStatusCode)
                return (false, $"HTTP {(int)res.StatusCode}\n\n{body}");

            return (true, string.IsNullOrWhiteSpace(body) ? "OK (empty response)" : body);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    private static string AppendQuery(string url, string key, string value)
    {
        var sep = url.Contains("?") ? "&" : "?";
        return url + sep + Uri.EscapeDataString(key) + "=" + Uri.EscapeDataString(value);
    }
}

internal static class CleanupScheduler
{
    private static System.Threading.Timer? _timer;
    private static int _running = 0;

    public static event EventHandler<(bool ok, string info)>? CleanupFinished;

    public static void ApplyFromSettings()
    {
        Stop();

        var s = SettingsStore.Current;

        if (!s.CleanupDailyEnabled && !s.CleanupEveryEnabled) return;

        if (s.CleanupDailyEnabled) ScheduleNextDaily();
        else if (s.CleanupEveryEnabled) ScheduleEveryMinutes();
    }

    public static void Stop()
    {
        var t = Interlocked.Exchange(ref _timer, null);
        t?.Dispose();
    }

    private static void ScheduleNextDaily()
    {
        var s = SettingsStore.Current;

        var now = DateTime.Now;
        var next = new DateTime(now.Year, now.Month, now.Day, s.CleanupDailyHour, s.CleanupDailyMinute, 0);
        if (next <= now) next = next.AddDays(1);

        var due = next - now;
        if (due < TimeSpan.Zero) due = TimeSpan.Zero;

        _timer = new System.Threading.Timer(_ => { _ = RunCleanupOnceAsync(rescheduleDaily: true); },
            null, due, Timeout.InfiniteTimeSpan);
    }

    private static void ScheduleEveryMinutes()
    {
        var s = SettingsStore.Current;

        var minutes = s.CleanupEveryMinutes;
        if (minutes < 1) minutes = 1;

        var period = TimeSpan.FromMinutes(minutes);

        _timer = new System.Threading.Timer(_ => { _ = RunCleanupOnceAsync(rescheduleDaily: false); },
            null, period, period);
    }

    private static async Task RunCleanupOnceAsync(bool rescheduleDaily)
    {
        if (Interlocked.Exchange(ref _running, 1) == 1) return;

        try
        {
            var (ok, info) = await CleanupApi.DeleteInboxAllAsync();
            CleanupFinished?.Invoke(null, (ok, info));
        }
        catch (Exception ex)
        {
            CleanupFinished?.Invoke(null, (false, ex.Message));
        }
        finally
        {
            Interlocked.Exchange(ref _running, 0);

            if (rescheduleDaily)
            {
                var s = SettingsStore.Current;
                if (s.CleanupDailyEnabled && !s.CleanupEveryEnabled) ScheduleNextDaily();
                else Stop();
            }
        }
    }
}

internal static class I18n
{
    private static readonly object _lock = new();
    private static string _lang = "en";

    private static readonly Dictionary<string, Dictionary<string, string>> _dict =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["en"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["TrayTitle"] = "Clipboard → URL Sender",
                ["MenuEnabled"] = "Enabled",
                ["MenuDisabled"] = "Disabled",
                ["MenuDeleteInbox"] = "Delete all INBOX notes",
                ["MenuSettings"] = "Settings...",
                ["MenuHelp"] = "Help...",
                ["MenuAbout"] = "About...",
                ["MenuExit"] = "Exit",

                ["ConfirmTitle"] = "Confirm",
                ["ConfirmDeleteInboxBody"] = "This will delete ALL notes in INBOX.\nDo you want to continue?",
                ["CleanupTitle"] = "Cleanup_API",
                ["CleanupRunning"] = "Deleting INBOX notes...",
                ["CleanupDoneTitle"] = "Delete INBOX: Done",
                ["CleanupFailTitle"] = "Delete INBOX: Failed",

                ["SendOkTitle"] = "Send OK",
                ["SendFailTitle"] = "Send Failed",

                ["BalloonEnabled"] = "Sending enabled",
                ["BalloonDisabled"] = "Sending disabled",

                ["AutoCleanupTitle"] = "Cleanup_API (Auto)",
                ["AutoCleanupOk"] = "Delete INBOX: Success",
                ["AutoCleanupFail"] = "Delete INBOX: Failed",

                ["CleanupModeInfoFormat"] = "(Auto-delete mode: {0})",
                ["CleanupModeOff"] = "OFF",
                ["CleanupModeDailyFormat"] = "Daily {0:00}:{1:00}",
                ["CleanupModeEveryFormat"] = "Every {0} min",
                ["CleanupModeDailyAndEveryFormat"] = "Daily {0:00}:{1:00} + Every {2} min (both ON)",

                ["CleanupLastNotYet"] = "(Last auto-delete: not yet)",
                ["CleanupLastFormat"] = "(Last auto-delete: {0} / {1})",

                ["WordSuccess"] = "Success",
                ["WordFail"] = "Fail",

                // About
                ["AboutTitle"] = "About",
                ["AboutBodyFormat"] = "{0}\nVersion: {1}\n\nHelp: {2}",

                // SettingsForm (省略せずそのまま維持)
                ["SettingsTitle"] = "Settings",
                ["LangLabel"] = "Language",
                ["LangEnglish"] = "English",
                ["LangJapanese"] = "日本語",
                ["LangTurkish"] = "Türkçe",
                ["Save"] = "Save",
                ["Cancel"] = "Cancel",
                ["ShowToken"] = "Show token",
                ["ShowCleanupToken"] = "Show Cleanup token",
                ["TestConnection"] = "Test",
                ["EnabledCheckbox"] = "Enable sending",
                ["ShowSuccessCheckbox"] = "Show message on success",
                ["HotkeyLabel"] = "Hotkey (click and press)",
                ["PostUrlLabel"] = "POST URL",
                ["TokenLabel"] = "Token",
                ["CleanupSection"] = "Cleanup_API (Delete INBOX)",
                ["CleanupUrlLabel"] = "Cleanup_API URL",
                ["CleanupTokenLabel"] = "Cleanup_API token",
                ["DailyCheckbox"] = "Run INBOX delete daily (once per day)",
                ["EveryCheckbox"] = "Run INBOX delete every X minutes",
                ["TimeLabel"] = "Time",
                ["Hour"] = "hour",
                ["Minute"] = "min",
                ["EveryMinutesLabel"] = "Interval (minutes)",

                ["SavedMsg"] = "Saved",
                ["SavedTitle"] = "OK",
                ["InputErrorTitle"] = "Input error",
                ["UrlInvalid"] = "URL is invalid.\nExample: https://sample.com/notemod/ta/api/api.php",
                ["TokenEmpty"] = "Token is empty",
                ["NeedModifier"] = "Use Ctrl or Alt (or Shift) with another key",

                ["BasicAuthSection"] = "Basic Auth (optional)",
                ["BasicUserLabel"] = "Username",
                ["BasicPassLabel"] = "Password",
                ["ShowBasicPass"] = "Show password",
                ["BasicIncomplete"] = "If you set a username, please also set a password",
            },

            ["ja"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["TrayTitle"] = "Clipboard → URL Sender",
                ["MenuEnabled"] = "有効",
                ["MenuDisabled"] = "無効",
                ["MenuDeleteInbox"] = "INBOX全削除",
                ["MenuSettings"] = "設定",
                ["MenuHelp"] = "ヘルプ",
                ["MenuAbout"] = "バージョン情報",
                ["MenuExit"] = "終了",

                ["ConfirmTitle"] = "最終確認",
                ["ConfirmDeleteInboxBody"] = "INBOX内のノートを全て削除します\n本当に実行しますか？",
                ["CleanupTitle"] = "Cleanup_API",
                ["CleanupRunning"] = "INBOX全削除を実行中...",
                ["CleanupDoneTitle"] = "INBOX全削除：完了",
                ["CleanupFailTitle"] = "INBOX全削除：失敗",

                ["SendOkTitle"] = "送信OK",
                ["SendFailTitle"] = "送信失敗",

                ["BalloonEnabled"] = "送信を有効化しました",
                ["BalloonDisabled"] = "送信を無効化しました",

                ["AutoCleanupTitle"] = "Cleanup_API（自動）",
                ["AutoCleanupOk"] = "INBOX全削除：成功",
                ["AutoCleanupFail"] = "INBOX全削除：失敗",

                ["CleanupModeInfoFormat"] = "(自動削除モード: {0})",
                ["CleanupModeOff"] = "OFF",
                ["CleanupModeDailyFormat"] = "定時 {0:00}:{1:00}",
                ["CleanupModeEveryFormat"] = "{0}分毎",
                ["CleanupModeDailyAndEveryFormat"] = "定時 {0:00}:{1:00} + {2}分毎（※両方ON）",

                ["CleanupLastNotYet"] = "(最終自動削除: 未実行)",
                ["CleanupLastFormat"] = "(最終自動削除: {0} / {1})",

                ["WordSuccess"] = "成功",
                ["WordFail"] = "失敗",

                // About
                ["AboutTitle"] = "バージョン情報",
                ["AboutBodyFormat"] = "{0}\nVersion: {1}\n\nヘルプ: {2}",

                // SettingsForm
                ["SettingsTitle"] = "設定",
                ["LangLabel"] = "言語",
                ["LangEnglish"] = "English",
                ["LangJapanese"] = "日本語",
                ["LangTurkish"] = "Türkçe",
                ["Save"] = "保存",
                ["Cancel"] = "キャンセル",
                ["ShowToken"] = "token を表示する",
                ["ShowCleanupToken"] = "Cleanup token を表示する",
                ["TestConnection"] = "接続テスト",
                ["EnabledCheckbox"] = "送信を有効にする",
                ["ShowSuccessCheckbox"] = "成功時もメッセージを表示する",
                ["HotkeyLabel"] = "ホットキー（クリックして押す）",
                ["PostUrlLabel"] = "POST先URL",
                ["TokenLabel"] = "token",
                ["CleanupSection"] = "Cleanup_API（INBOX全削除）",
                ["CleanupUrlLabel"] = "Cleanup_API URL",
                ["CleanupTokenLabel"] = "Cleanup_API token",
                ["DailyCheckbox"] = "定時でINBOX全削除を実行する（1日1回）",
                ["EveryCheckbox"] = "x分毎にINBOX全削除を実行する",
                ["TimeLabel"] = "時刻",
                ["Hour"] = "時",
                ["Minute"] = "分",
                ["EveryMinutesLabel"] = "間隔（分）",

                ["SavedMsg"] = "保存しました",
                ["SavedTitle"] = "OK",
                ["InputErrorTitle"] = "入力エラー",
                ["UrlInvalid"] = "URLが正しくないです\n例: https://sample.com/notemod/ta/api/api.php",
                ["TokenEmpty"] = "tokenが空です",
                ["NeedModifier"] = "Ctrl または Alt（または Shift）と組み合わせてください",

                ["BasicAuthSection"] = "BASIC認証（任意）",
                ["BasicUserLabel"] = "ユーザー名",
                ["BasicPassLabel"] = "パスワード",
                ["ShowBasicPass"] = "パスワードを表示する",
                ["BasicIncomplete"] = "ユーザー名を設定した場合は、パスワードも設定してください",
            },

            ["tr"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["TrayTitle"] = "Clipboard → URL Sender",
                ["MenuEnabled"] = "Etkin",
                ["MenuDisabled"] = "Devre dışı",
                ["MenuDeleteInbox"] = "INBOX'taki tüm notları sil",
                ["MenuSettings"] = "Ayarlar...",
                ["MenuHelp"] = "Yardım...",
                ["MenuAbout"] = "Hakkında...",
                ["MenuExit"] = "Çıkış",

                ["ConfirmTitle"] = "Onay",
                ["ConfirmDeleteInboxBody"] = "Bu işlem INBOX içindeki TÜM notları silecektir.\nDevam etmek istiyor musunuz?",
                ["CleanupTitle"] = "Cleanup_API",
                ["CleanupRunning"] = "INBOX notları siliniyor...",
                ["CleanupDoneTitle"] = "INBOX silme: Tamamlandı",
                ["CleanupFailTitle"] = "INBOX silme: Başarısız",

                ["SendOkTitle"] = "Gönderme başarılı",
                ["SendFailTitle"] = "Gönderme başarısız",

                ["BalloonEnabled"] = "Gönderme etkin",
                ["BalloonDisabled"] = "Gönderme devre dışı",

                ["AutoCleanupTitle"] = "Cleanup_API (Otomatik)",
                ["AutoCleanupOk"] = "INBOX silme: Başarılı",
                ["AutoCleanupFail"] = "INBOX silme: Başarısız",

                ["CleanupModeInfoFormat"] = "(Otomatik silme modu: {0})",
                ["CleanupModeOff"] = "KAPALI",
                ["CleanupModeDailyFormat"] = "Günlük {0:00}:{1:00}",
                ["CleanupModeEveryFormat"] = "{0} dakikada bir",
                ["CleanupModeDailyAndEveryFormat"] = "Günlük {0:00}:{1:00} + {2} dakikada bir (ikisi de açık)",

                ["CleanupLastNotYet"] = "(Son otomatik silme: henüz yok)",
                ["CleanupLastFormat"] = "(Son otomatik silme: {0} / {1})",

                ["WordSuccess"] = "Başarılı",
                ["WordFail"] = "Başarısız",

                // About
                ["AboutTitle"] = "Hakkında",
                ["AboutBodyFormat"] = "{0}\nSürüm: {1}\n\nYardım: {2}",

                // SettingsForm
                ["SettingsTitle"] = "Ayarlar",
                ["LangLabel"] = "Dil",
                ["LangEnglish"] = "English",
                ["LangJapanese"] = "日本語",
                ["LangTurkish"] = "Türkçe",
                ["Save"] = "Kaydet",
                ["Cancel"] = "İptal",
                ["ShowToken"] = "Token'ı göster",
                ["ShowCleanupToken"] = "Cleanup token'ını göster",
                ["TestConnection"] = "Test",
                ["EnabledCheckbox"] = "Göndermeyi etkinleştir",
                ["ShowSuccessCheckbox"] = "Başarıda da mesaj göster",
                ["HotkeyLabel"] = "Kısayol (tıkla ve bas)",
                ["PostUrlLabel"] = "POST URL",
                ["TokenLabel"] = "Token",
                ["CleanupSection"] = "Cleanup_API (INBOX sil)",
                ["CleanupUrlLabel"] = "Cleanup_API URL",
                ["CleanupTokenLabel"] = "Cleanup_API token",
                ["DailyCheckbox"] = "INBOX silmeyi günlük çalıştır (günde 1 kez)",
                ["EveryCheckbox"] = "INBOX silmeyi X dakikada bir çalıştır",
                ["TimeLabel"] = "Saat",
                ["Hour"] = "saat",
                ["Minute"] = "dk",
                ["EveryMinutesLabel"] = "Aralık (dakika)",

                ["SavedMsg"] = "Kaydedildi",
                ["SavedTitle"] = "OK",
                ["InputErrorTitle"] = "Girdi hatası",
                ["UrlInvalid"] = "URL geçersiz.\nÖrnek: https://sample.com/notemod/ta/api/api.php",
                ["TokenEmpty"] = "Token boş",
                ["NeedModifier"] = "Ctrl veya Alt (ya da Shift) + başka bir tuş kullanın",

                ["BasicAuthSection"] = "Basic Auth (isteğe bağlı)",
                ["BasicUserLabel"] = "Kullanıcı adı",
                ["BasicPassLabel"] = "Parola",
                ["ShowBasicPass"] = "Parolayı göster",
                ["BasicIncomplete"] = "Kullanıcı adı girerseniz parolayı da girin",
            },
        };

    public static IReadOnlyList<(string code, string displayName)> SupportedLanguages { get; } =
        new List<(string, string)>
        {
            ("en", "English"),
            ("ja", "日本語"),
            ("tr", "Türkçe"),
        };

    public static void SetLanguage(string? langCode)
    {
        lock (_lock)
        {
            if (string.IsNullOrWhiteSpace(langCode)) { _lang = "en"; return; }
            _lang = _dict.ContainsKey(langCode) ? langCode : "en";
        }
    }

    public static string CurrentLanguage { get { lock (_lock) return _lang; } }

    public static string T(string key)
    {
        lock (_lock)
        {
            if (_dict.TryGetValue(_lang, out var d) && d.TryGetValue(key, out var v)) return v;
            if (_dict.TryGetValue("en", out var en) && en.TryGetValue(key, out var ev)) return ev;
            return key;
        }
    }
}