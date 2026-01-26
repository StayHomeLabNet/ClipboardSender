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
using System.Diagnostics;
using System.Linq;

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
    // ======= 送信（既存） =======
    public string BaseUrl { get; set; } = "";
    public string TokenEncrypted { get; set; } = "";

    public bool Enabled { get; set; } = true;
    public bool ShowMessageOnSuccess { get; set; } = true;

    public uint HotkeyModifiers { get; set; } = 0x0001 | 0x0002; // Alt|Ctrl
    public int HotkeyVk { get; set; } = 0x4E; // N
    public string HotkeyDisplay { get; set; } = "Ctrl + Alt + N";

    // ======= Cleanup（既存） =======
    public string CleanupBaseUrl { get; set; } = "";
    public string CleanupTokenEncrypted { get; set; } = "";
    public bool CleanupPretty { get; set; } = true;

    public bool CleanupDailyEnabled { get; set; } = false;
    public int CleanupDailyHour { get; set; } = 2;
    public int CleanupDailyMinute { get; set; } = 0;

    public bool CleanupEveryEnabled { get; set; } = false;
    public int CleanupEveryMinutes { get; set; } = 60;

    // ======= 受信（追加） =======
    public string ReceiveBaseUrl { get; set; } = "";
    public uint ReceiveHotkeyModifiers { get; set; } = 0x0001 | 0x0002; // Alt|Ctrl
    public int ReceiveHotkeyVk { get; set; } = 0x52; // R
    public string ReceiveHotkeyDisplay { get; set; } = "Ctrl + Alt + R";

    // ★追加：受信でペーストまで行うか
    public bool ReceiveAutoPaste { get; set; } = true;

    // ★追加：「クリップボード安定待ち」ms
    public int ClipboardStableWaitMs { get; set; } = 60;

    // ======= UI =======
    public string Language { get; set; } = "en"; // "en" / "ja" / "tr"

    // BASIC（共通）
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
                var s = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();

                // ★互換：過去に別名で保存されていても読めるようにする
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                // 別名: ReceivePasteEnabled -> ReceiveAutoPaste
                if (root.TryGetProperty("ReceivePasteEnabled", out var p1) &&
                    (p1.ValueKind == JsonValueKind.True || p1.ValueKind == JsonValueKind.False))
                {
                    s.ReceiveAutoPaste = p1.GetBoolean();
                }

                // 別名: ReceiveClipboardStabilizeWaitMs -> ClipboardStableWaitMs
                if (root.TryGetProperty("ReceiveClipboardStabilizeWaitMs", out var p2) &&
                    p2.ValueKind == JsonValueKind.Number)
                {
                    s.ClipboardStableWaitMs = p2.GetInt32();
                }

                return s;
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
                    using var ico = new System.Drawing.Icon(stream);
                    return (System.Drawing.Icon)ico.Clone();
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
    private const string HELP_URL = "https://stayhomelab.net/ClipboardSync";

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

        // ★言語変更が「即反映」されるようにする（保存前でも反映）
        I18n.LanguageChanged += (_, __) =>
        {
            RunOnUi(() =>
            {
                RefreshTrayTexts();
                SyncEnabledMenu();
                UpdateCleanupModeInfo();
                UpdateCleanupLastResultInfo();
            });
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
        _menuHelp = new ToolStripMenuItem();
        _menuAbout = new ToolStripMenuItem();
        _menuExit = new ToolStripMenuItem();

        _tray.ContextMenuStrip.Items.Add(_enabledMenuItem);
        _tray.ContextMenuStrip.Items.Add(new ToolStripSeparator());

        _tray.ContextMenuStrip.Items.Add(_cleanupModeInfoItem);
        _tray.ContextMenuStrip.Items.Add(_cleanupLastResultItem);
        _tray.ContextMenuStrip.Items.Add(new ToolStripSeparator());

        _tray.ContextMenuStrip.Items.Add(_menuDeleteInbox);
        _tray.ContextMenuStrip.Items.Add(_menuSettings);

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

        // ★送信（Clipboard → API）
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

        // ★受信（Hotkey → Read API → Clipboard → (Option) Paste）
        _watcher.ReceiveHotkeyPressed += async (_, __) =>
        {
            var (ok, textOrErr) = await Receiver.FetchLatestNoteTextAsync();

            if (!ok)
            {
                using var top = new Form { TopMost = true, ShowInTaskbar = false };
                top.StartPosition = FormStartPosition.Manual;
                top.Location = new System.Drawing.Point(-2000, -2000);
                top.Show();

                MessageBox.Show(top, textOrErr, I18n.T("ReceiveFailTitle"),
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var text = textOrErr ?? "";
            if (string.IsNullOrWhiteSpace(text))
            {
                _tray.BalloonTipTitle = I18n.T("ReceiveTitle");
                _tray.BalloonTipText = I18n.T("ReceiveEmpty");
                _tray.ShowBalloonTip(1000);
                return;
            }

            await ClipboardUtil.TrySetTextAsync(text);

            // ★「クリップボード安定待ち」
            var waitMs = Math.Max(0, SettingsStore.Current.ClipboardStableWaitMs);
            if (waitMs > 0) await Task.Delay(waitMs);

            // ★チェックありならペーストまで実行
            if (SettingsStore.Current.ReceiveAutoPaste)
            {
                var okPaste = PasteHelper.CtrlV_FallbackSendKeys();
                if (!okPaste)
                {
                    _tray.BalloonTipTitle = I18n.T("ReceiveTitle");
                    _tray.BalloonTipText =
                        "Auto paste failed.\n" +
                        "・貼り付け先が管理者権限のアプリだと拒否されます\n" +
                        "・Notepadで試してもダメなら設定/実装側の問題です";
                    _tray.ShowBalloonTip(2000);
                }
            }

            _tray.BalloonTipTitle = I18n.T("ReceiveTitle");
            _tray.BalloonTipText = I18n.T("ReceiveOkBalloon");
            _tray.ShowBalloonTip(1000);
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
        var asm = Assembly.GetExecutingAssembly();
        var prod = asm.GetCustomAttributes<AssemblyProductAttribute>().FirstOrDefault()?.Product;
        if (!string.IsNullOrWhiteSpace(prod)) return prod!;
        return asm.GetName().Name ?? "Clipboard Sync";
    }

    public static string GetVersionString()
    {
        var asm = Assembly.GetExecutingAssembly();
        var info = asm.GetCustomAttributes<AssemblyInformationalVersionAttribute>()
                      .FirstOrDefault()?.InformationalVersion;

        if (!string.IsNullOrWhiteSpace(info)) return info!;

        var fv = asm.GetCustomAttributes<AssemblyFileVersionAttribute>()
                    .FirstOrDefault()?.Version;
        if (!string.IsNullOrWhiteSpace(fv)) return fv!;

        return asm.GetName().Version?.ToString() ?? "0.0.0";
    }
}

internal sealed class ClipboardWatcherForm : Form
{
    private const int WM_HOTKEY = 0x0312;

    private const int HOTKEY_ID_TOGGLE = 1;
    private const int HOTKEY_ID_RECEIVE = 2;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, int vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public event EventHandler<string>? ClipboardTextCopied;
    public event EventHandler<bool>? EnabledToggled;
    public event EventHandler? ReceiveHotkeyPressed;

    private string? _lastClipboardText;
    private DateTime _lastEventAtUtc = DateTime.MinValue;

    private static readonly TimeSpan Debounce = TimeSpan.FromMilliseconds(350);

    private bool _registeredToggle = false;
    private bool _registeredReceive = false;

    public void ApplyHotkeyFromSettings()
    {
        if (_registeredToggle)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID_TOGGLE);
            _registeredToggle = false;
        }
        if (_registeredReceive)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID_RECEIVE);
            _registeredReceive = false;
        }

        var s = SettingsStore.Current;

        _registeredToggle = RegisterHotKey(this.Handle, HOTKEY_ID_TOGGLE, s.HotkeyModifiers, s.HotkeyVk);
        _registeredReceive = RegisterHotKey(this.Handle, HOTKEY_ID_RECEIVE, s.ReceiveHotkeyModifiers, s.ReceiveHotkeyVk);
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
        else if (m.Msg == WM_HOTKEY)
        {
            var id = m.WParam.ToInt32();
            if (id == HOTKEY_ID_TOGGLE) ToggleEnabled();
            else if (id == HOTKEY_ID_RECEIVE) ReceiveHotkeyPressed?.Invoke(this, EventArgs.Empty);
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
        UnregisterHotKey(this.Handle, HOTKEY_ID_RECEIVE);
        RemoveClipboardFormatListener(this.Handle);
        base.OnHandleDestroyed(e);
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool AddClipboardFormatListener(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
}

internal static class ClipboardUtil
{
    public static async Task TrySetTextAsync(string text)
    {
        text = (text ?? "").Replace("\r\n", "\n");

        for (int i = 0; i < 5; i++)
        {
            try
            {
                Clipboard.SetText(text, TextDataFormat.UnicodeText);
                return;
            }
            catch
            {
                await Task.Delay(40);
            }
        }
    }
}

internal static class Receiver
{
    private static readonly HttpClient Client = new()
    {
        Timeout = TimeSpan.FromSeconds(15)
    };

    public static async Task<(bool ok, string infoOrText)> FetchLatestNoteTextAsync()
    {
        try
        {
            var s = SettingsStore.Current;

            var baseUrl = (s.ReceiveBaseUrl ?? "").Trim();
            if (string.IsNullOrWhiteSpace(baseUrl))
                return (false, "Receive (Read API) URL is empty");

            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var _))
                return (false, "Receive (Read API) URL is invalid");

            var token = (DpapiHelper.Decrypt(s.TokenEncrypted) ?? "").Trim();
            if (string.IsNullOrWhiteSpace(token))
                return (false, "Token is empty or cannot be decrypted");

            var readUrl = NormalizeReadApiUrl(baseUrl);

            var url = AppendQuery(readUrl, "token", token);
            url = AppendQuery(url, "action", "latest_note");

            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            BasicAuth.Apply(req, s);

            using var res = await Client.SendAsync(req);
            var body = await res.Content.ReadAsStringAsync();

            if (!res.IsSuccessStatusCode)
                return (false, $"HTTP {(int)res.StatusCode}\n\n{body}");

            return (true, body ?? "");
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    private static string NormalizeReadApiUrl(string url)
    {
        if (url.EndsWith("/api.php", StringComparison.OrdinalIgnoreCase))
            return url.Substring(0, url.Length - "/api.php".Length) + "/read_api.php";

        if (url.EndsWith("/api/", StringComparison.OrdinalIgnoreCase))
            return url + "read_api.php";

        if (url.EndsWith("/api", StringComparison.OrdinalIgnoreCase))
            return url + "/read_api.php";

        return url;
    }

    private static string AppendQuery(string url, string key, string value)
    {
        var sep = url.Contains("?") ? "&" : "?";
        return url + sep + Uri.EscapeDataString(key) + "=" + Uri.EscapeDataString(value);
    }
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

            var token = (DpapiHelper.Decrypt(s.TokenEncrypted) ?? "").Trim();
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

            var token = (DpapiHelper.Decrypt(s.CleanupTokenEncrypted) ?? "").Trim();
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

    // バックアップファイル数（purge_bak=1 & dry_run=2 で "count" を取得）
    public static async Task<(bool ok, int count, string info)> GetBackupCountAsync()
    {
        try
        {
            var s = SettingsStore.Current;

            var baseUrl = (s.CleanupBaseUrl ?? "").Trim();
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var _))
                return (false, -1, "Cleanup_API URL is not set or invalid");

            var token = (DpapiHelper.Decrypt(s.CleanupTokenEncrypted) ?? "").Trim();
            if (string.IsNullOrWhiteSpace(token))
                return (false, -1, "Cleanup_API token is not set / cannot be decrypted");

            var url = baseUrl;
            if (s.CleanupPretty) url = AppendQuery(baseUrl, "pretty", "1");

            var kv = new List<KeyValuePair<string, string>>
            {
                new("token", token),
                new("purge_bak", "1"),
                new("dry_run", "2"),
            };

            using var content = new FormUrlEncodedContent(kv);
            content.Headers.ContentType!.CharSet = "utf-8";

            using var req = new HttpRequestMessage(HttpMethod.Post, url) { Content = content };
            BasicAuth.Apply(req, s);

            using var res = await Client.SendAsync(req);
            var body = await res.Content.ReadAsStringAsync();

            if (!res.IsSuccessStatusCode)
                return (false, -1, $"HTTP {(int)res.StatusCode}\n\n{body}");

            var count = TryParseCount(body, out var parsed) ? parsed : -1;
            if (count < 0)
                return (false, -1, string.IsNullOrWhiteSpace(body) ? "count not found" : body);

            return (true, count, body ?? "");
        }
        catch (Exception ex)
        {
            return (false, -1, ex.Message);
        }
    }

    // バックアップファイル一括削除（purge_bak=1 & confirm=YES）
    public static async Task<(bool ok, string info)> PurgeBackupsAsync()
    {
        try
        {
            var s = SettingsStore.Current;

            var baseUrl = (s.CleanupBaseUrl ?? "").Trim();
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var _))
                return (false, "Cleanup_API URL is not set or invalid");

            var token = (DpapiHelper.Decrypt(s.CleanupTokenEncrypted) ?? "").Trim();
            if (string.IsNullOrWhiteSpace(token))
                return (false, "Cleanup_API token is not set / cannot be decrypted");

            var url = baseUrl;
            if (s.CleanupPretty) url = AppendQuery(baseUrl, "pretty", "1");

            var kv = new List<KeyValuePair<string, string>>
            {
                new("token", token),
                new("purge_bak", "1"),
                new("confirm", "YES"),
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

    private static bool TryParseCount(string? body, out int count)
    {
        count = -1;
        if (string.IsNullOrWhiteSpace(body)) return false;

        // まず JSON として読む
        try
        {
            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;

            // "count": 12
            if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("count", out var c) && c.ValueKind == JsonValueKind.Number)
            {
                count = c.GetInt32();
                return true;
            }

            // もし {"data":{"count":12}} とかでも拾えるように軽く探索
            foreach (var prop in root.EnumerateObject())
            {
                if (prop.Value.ValueKind == JsonValueKind.Object &&
                    prop.Value.TryGetProperty("count", out var cc) &&
                    cc.ValueKind == JsonValueKind.Number)
                {
                    count = cc.GetInt32();
                    return true;
                }
            }
        }
        catch
        {
            // JSONじゃない場合は下へ
        }

        // フォールバック: "count": の後ろの数字を拾う
        try
        {
            var idx = body.IndexOf("\"count\"", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return false;

            var colon = body.IndexOf(':', idx);
            if (colon < 0) return false;

            int i = colon + 1;
            while (i < body.Length && char.IsWhiteSpace(body[i])) i++;

            int start = i;
            while (i < body.Length && char.IsDigit(body[i])) i++;

            if (i <= start) return false;

            var num = body.Substring(start, i - start);
            if (int.TryParse(num, out var n))
            {
                count = n;
                return true;
            }
        }
        catch { }

        return false;
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

    // ★追加：言語変更通知（保存前でもUIが即反映できる）
    public static event EventHandler? LanguageChanged;

    private static readonly Dictionary<string, Dictionary<string, string>> _dict =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["en"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["TrayTitle"] = "Clipboard Sync",
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

                ["AboutTitle"] = "About",
                ["AboutBodyFormat"] = "{0}\nVersion: {1}\n\nHelp: {2}",

                ["ReceiveTitle"] = "Receive",
                ["ReceiveOkBalloon"] = "Latest note copied to clipboard",
                ["ReceiveEmpty"] = "Latest note is empty",
                ["ReceiveFailTitle"] = "Receive Failed",

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

                ["ReceiveUrlLabel"] = "Read API URL",
                ["ReceiveHotkeyLabel"] = "Receive hotkey (click and press)",
                ["TabSendDelete"] = "Send/Delete",
                ["TabReceive"] = "Receive",

                // ★追加：受信設定
                ["ReceiveAutoPaste"] = "Auto paste after receive",
                ["ClipboardStableWaitLabel"] = "Clipboard stabilize wait (ms)",

                // バックアップファイル表示/削除
                ["BakCountLoading"] = "Backup file: (loading...)",
                ["BakCountNone"] = "Backup file: -",
                ["BakCountFormat"] = "Backup file(s): {0}",
                ["BakCountFail"] = "Backup file: ?",
                ["PurgeBakButton"] = "Delete ALL backup files",
                ["ConfirmPurgeBakBody"] = "This will delete ALL backup files.\nDo you want to continue?",
                ["PurgeBakDoneTitle"] = "Backup file delete: Done",
                ["PurgeBakFailTitle"] = "Backup file delete: Failed",
            },

            ["ja"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["TrayTitle"] = "Clipboard Sync",
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

                ["AboutTitle"] = "バージョン情報",
                ["AboutBodyFormat"] = "{0}\nVersion: {1}\n\nヘルプ: {2}",

                ["ReceiveTitle"] = "受信",
                ["ReceiveOkBalloon"] = "最新ノートをクリップボードへコピーしました",
                ["ReceiveEmpty"] = "最新ノートが空です",
                ["ReceiveFailTitle"] = "受信失敗",

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

                ["ReceiveUrlLabel"] = "Read API URL（受信）",
                ["ReceiveHotkeyLabel"] = "受信ホットキー（クリックして押す）",
                ["TabSendDelete"] = "送信/削除設定",
                ["TabReceive"] = "受信設定",

                // ★追加：受信設定
                ["ReceiveAutoPaste"] = "ペーストまで含める",
                ["ClipboardStableWaitLabel"] = "クリップボード安定待ち（ms）",

                // バックアップファイル表示/削除
                ["BakCountLoading"] = "バックアップファイル: 取得中...",
                ["BakCountNone"] = "バックアップファイル: -",
                ["BakCountFormat"] = "バックアップファイル: {0}",
                ["BakCountFail"] = "バックアップファイル: ?",
                ["PurgeBakButton"] = "バックアップファイルを全削除",
                ["ConfirmPurgeBakBody"] = "バックアップファイルを全て削除します\n本当に実行しますか？",
                ["PurgeBakDoneTitle"] = "バックアップファイル削除：完了",
                ["PurgeBakFailTitle"] = "バックアップファイル削除：失敗",
            },

            ["tr"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["TrayTitle"] = "Clipboard Sync",
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

                ["AboutTitle"] = "Hakkında",
                ["AboutBodyFormat"] = "{0}\nSürüm: {1}\n\nYardım: {2}",

                ["ReceiveTitle"] = "Al",
                ["ReceiveOkBalloon"] = "En son not panoya kopyalandı",
                ["ReceiveEmpty"] = "En son not boş",
                ["ReceiveFailTitle"] = "Alma başarısız",

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

                ["ReceiveUrlLabel"] = "Read API URL",
                ["ReceiveHotkeyLabel"] = "Alma kısayolu (tıkla ve bas)",
                ["TabSendDelete"] = "Gönder/Sil",
                ["TabReceive"] = "Al",

                // 受信設定
                ["ReceiveAutoPaste"] = "Alınca otomatik yapıştır",
                ["ClipboardStableWaitLabel"] = "Pano stabilize bekleme (ms)",

                // バックアップ表示/削除
                ["BakCountLoading"] = "Yedek dosyası: (yükleniyor...)",
                ["BakCountNone"] = "Yedek dosyası: -",
                ["BakCountFormat"] = "Yedek dosyası: {0}",
                ["BakCountFail"] = "Yedek dosyası: ?",
                ["PurgeBakButton"] = "Tüm yedek dosyalarını sil",
                ["ConfirmPurgeBakBody"] = "TÜM yedek dosyaları silinecek.\nDevam etmek istiyor musunuz?",
                ["PurgeBakDoneTitle"] = "Yedek dosyası silme: Tamam",
                ["PurgeBakFailTitle"] = "Yedek dosyası silme: Başarısız",
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
        bool changed = false;

        lock (_lock)
        {
            var before = _lang;

            if (string.IsNullOrWhiteSpace(langCode)) _lang = "en";
            else _lang = _dict.ContainsKey(langCode) ? langCode : "en";

            changed = !string.Equals(before, _lang, StringComparison.OrdinalIgnoreCase);
        }

        if (changed) LanguageChanged?.Invoke(null, EventArgs.Empty);
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

internal static class PasteHelper
{
    private const int INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_SCANCODE = 0x0008;

    private const ushort VK_CONTROL = 0x11;
    private const ushort VK_V = 0x56;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern uint MapVirtualKey(uint uCode, uint uMapType);

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public int type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private static INPUT ScanDown(ushort vk)
    {
        var scan = (ushort)MapVirtualKey(vk, 0); // MAPVK_VK_TO_VSC
        return new INPUT
        {
            type = INPUT_KEYBOARD,
            U = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = scan,
                    dwFlags = KEYEVENTF_SCANCODE,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero
                }
            }
        };
    }

    private static INPUT ScanUp(ushort vk)
    {
        var scan = (ushort)MapVirtualKey(vk, 0);
        return new INPUT
        {
            type = INPUT_KEYBOARD,
            U = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = scan,
                    dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero
                }
            }
        };
    }

    /// <summary>
    /// Ctrl+V を実行。SendInputが失敗したら SendKeys にフォールバック。
    /// 戻り値: true=何かしら送れた / false=完全失敗
    /// </summary>
    public static bool CtrlV_FallbackSendKeys()
    {
        // まずは SendInput（スキャンコード）
        try
        {
            var inputs = new INPUT[]
            {
                ScanDown(VK_CONTROL),
                ScanDown(VK_V),
                ScanUp(VK_V),
                ScanUp(VK_CONTROL),
            };

            var sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
            if (sent == inputs.Length) return true;
        }
        catch { }

        // フォールバック: SendKeys（WinForms/STA前提）
        try
        {
            SendKeys.SendWait("^v");
            return true;
        }
        catch
        {
            return false;
        }
    }
}