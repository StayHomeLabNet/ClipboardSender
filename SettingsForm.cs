// SettingsForm.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;        // Help/Website link
using System.Net.Http;
using System.Net.Http.Headers;   // BASIC
using System.Reflection;         // version/product
using System.Text;               // BASIC
using System.Windows.Forms;
using System.Threading.Tasks;

internal sealed class SettingsForm : Form
{
    // ★ここを書き換えて使ってね
    private const string HelpUrl = "https://stayhomelab.net/";
    private const string WebsiteUrl = "https://github.com/StayHomeLabNet/ClipboardSender"; // ★Version右のWebサイトリンク

    // MOD_* の値（Program.csと合わせる）
    private const uint MOD_ALT = 0x0001;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;

    private readonly TextBox _url = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
    private readonly TextBox _token = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
    private readonly CheckBox _enabled = new();
    private readonly CheckBox _showSuccess = new();

    private readonly TextBox _cleanupUrl = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
    private readonly TextBox _cleanupToken = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, UseSystemPasswordChar = true };

    private readonly TextBox _hotkeyBox = new()
    {
        ReadOnly = true,
        TabStop = true
    };

    // 言語
    private readonly ComboBox _lang = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };

    // 自動削除（定時）
    private readonly CheckBox _cleanupDailyEnabled = new();
    private readonly NumericUpDown _cleanupDailyHour = new() { Minimum = 0, Maximum = 23, Width = 60 };
    private readonly NumericUpDown _cleanupDailyMinute = new() { Minimum = 0, Maximum = 59, Width = 60 };

    // 自動削除（x分毎）
    private readonly CheckBox _cleanupEveryEnabled = new();
    private readonly NumericUpDown _cleanupEveryMinutes = new() { Minimum = 1, Maximum = 1440, Width = 80 };

    private uint _pendingHotkeyMods;
    private int _pendingHotkeyVk;
    private string _pendingHotkeyDisplay = "";

    // ラベル類（言語切替でTextを差し替えるため保持）
    private readonly Label _lblUrl = new() { AutoSize = true, Left = 12, Top = 20 };
    private readonly Label _lblToken = new() { AutoSize = true, Left = 12, Top = 75 };
    private readonly Label _lblHotkey = new() { AutoSize = true, Left = 12 };
    private readonly Label _lblCleanup = new() { AutoSize = true, Left = 12 };
    private readonly Label _lblCleanupUrl = new() { AutoSize = true, Left = 12 };
    private readonly Label _lblCleanupToken = new() { AutoSize = true, Left = 12 };
    private readonly Label _lblDailyTime = new() { AutoSize = true, Left = 30 };
    private readonly Label _lblH = new() { AutoSize = true };
    private readonly Label _lblM = new() { AutoSize = true };
    private readonly Label _lblEvery = new() { AutoSize = true, Left = 30 };
    private readonly Label _lblLang = new() { AutoSize = true, Left = 12 };

    private readonly CheckBox _chkShowToken = new() { AutoSize = true, Checked = false };
    private readonly CheckBox _chkShowCleanupToken = new() { AutoSize = true, Checked = false };

    private readonly Button _btnTest = new() { Width = 120, Height = 30 };
    private readonly Button _btnSave = new() { Width = 120, Height = 30 };
    private readonly Button _btnCancel = new() { Width = 120, Height = 30 };

    // =========================
    // BASIC認証 UI
    // =========================
    private readonly Label _lblBasic = new() { AutoSize = true, Left = 12 };
    private readonly Label _lblBasicUser = new() { AutoSize = true, Left = 12 };
    private readonly TextBox _basicUser = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
    private readonly Label _lblBasicPass = new() { AutoSize = true, Left = 12 };
    private readonly TextBox _basicPass = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, UseSystemPasswordChar = true };
    private readonly CheckBox _chkShowBasicPass = new() { AutoSize = true, Checked = false };

    // =========================
    // ★最下段：Help / AppName / Version / Website
    // （Left指定をやめてFlowLayoutで一行配置）
    // =========================
    private readonly FlowLayoutPanel _bottomRow = new()
    {
        Dock = DockStyle.Bottom,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        FlowDirection = FlowDirection.LeftToRight,
        WrapContents = false,
        Padding = new Padding(12, 8, 12, 10)
    };

    private readonly Button _btnHelp = new() { AutoSize = true, Height = 26 };
    private readonly Label _lblAppName = new() { AutoSize = true };
    private readonly Label _lblVersion = new() { AutoSize = true };
    private readonly LinkLabel _lnkWebsite = new() { AutoSize = true };

    public SettingsForm()
    {
        Width = 560;
        Height = 950; // BASIC欄込み
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        // 位置は従来のまま
        _url.Top = 40; _url.Left = 12; _url.Width = 520;
        _token.Top = 95; _token.Left = 12; _token.Width = 520;

        _chkShowToken.Top = _token.Bottom + 5;
        _chkShowToken.Left = _token.Left;

        _showSuccess.Top = _chkShowToken.Bottom + 16;
        _showSuccess.Left = 12;
        _showSuccess.AutoSize = true;

        _btnTest.Left = 412;
        _btnTest.Top = _token.Bottom + 5;

        // BASIC欄（右側）
        _lblBasic.Top = _showSuccess.Top; _lblBasic.Left = 285;
        _lblBasicUser.Top = _lblBasic.Bottom + 6; _lblBasicUser.Left = 285;
        _basicUser.Top = _lblBasicUser.Bottom + 6; _basicUser.Left = 285; _basicUser.Width = 245;

        _lblBasicPass.Top = _basicUser.Bottom + 10; _lblBasicPass.Left = 285;
        _basicPass.Top = _lblBasicPass.Bottom + 6; _basicPass.Left = 285; _basicPass.Width = 245;

        _chkShowBasicPass.Top = _basicPass.Bottom + 6;
        _chkShowBasicPass.Left = 285;

        _enabled.Top = _basicUser.Top;
        _enabled.Left = 12;
        _enabled.AutoSize = true;

        _lblHotkey.Top = _lblBasicPass.Top;
        _hotkeyBox.Top = _lblHotkey.Bottom + 6;
        _hotkeyBox.Left = 12;
        _hotkeyBox.Width = 240;

        // Cleanup
        _lblCleanup.Top = _chkShowBasicPass.Top + 22;

        _lblCleanupUrl.Top = _lblCleanup.Bottom + 6;
        _cleanupUrl.Top = _lblCleanupUrl.Bottom + 6; _cleanupUrl.Left = 12; _cleanupUrl.Width = 520;

        _lblCleanupToken.Top = _cleanupUrl.Bottom + 10;
        _cleanupToken.Top = _lblCleanupToken.Bottom + 6; _cleanupToken.Left = 12; _cleanupToken.Width = 520;

        _chkShowCleanupToken.Top = _cleanupToken.Bottom + 5;
        _chkShowCleanupToken.Left = _cleanupToken.Left;

        // 自動削除UI
        _cleanupDailyEnabled.Top = _chkShowCleanupToken.Bottom + 15;
        _cleanupDailyEnabled.Left = 12;
        _cleanupDailyEnabled.AutoSize = true;

        _lblDailyTime.Top = _cleanupDailyEnabled.Bottom + 8;
        _cleanupDailyHour.Top = _lblDailyTime.Bottom + 4; _cleanupDailyHour.Left = 30;
        _lblH.Top = _cleanupDailyHour.Top + 4; _lblH.Left = _cleanupDailyHour.Right + 6;
        _cleanupDailyMinute.Top = _cleanupDailyHour.Top; _cleanupDailyMinute.Left = _lblH.Right + 10;
        _lblM.Top = _cleanupDailyMinute.Top + 4; _lblM.Left = _cleanupDailyMinute.Right + 6;

        _cleanupEveryEnabled.Top = _cleanupDailyHour.Bottom + 14;
        _cleanupEveryEnabled.Left = 12;
        _cleanupEveryEnabled.AutoSize = true;

        _lblEvery.Top = _cleanupEveryEnabled.Bottom + 8;
        _cleanupEveryMinutes.Top = _lblEvery.Bottom + 4; _cleanupEveryMinutes.Left = 30;

        // 言語
        _lblLang.Top = 770;
        _lang.Left = 12;
        _lang.Top = 800;

        _btnSave.Left = 412; _btnSave.Top = 800;
        _btnCancel.Left = 284; _btnCancel.Top = 800;

        // ★最下段：FlowLayoutで一行
        _bottomRow.Controls.Add(_btnHelp);

        // 間を少し空ける（見た目調整）
        _lblAppName.Margin = new Padding(14, 6, 0, 0);
        _lblVersion.Margin = new Padding(10, 6, 0, 0);
        _lnkWebsite.Margin = new Padding(10, 6, 0, 0);

        _bottomRow.Controls.Add(_lblAppName);
        _bottomRow.Controls.Add(_lblVersion);
        _bottomRow.Controls.Add(_lnkWebsite);

        // bottom row は Dock=Bottom なので Controls の最後に足す
        Controls.AddRange(new Control[]
        {
            _lblUrl, _url,
            _lblToken, _token,
            _chkShowToken,
            _btnTest,
            _showSuccess,

            // BASIC
            _lblBasic,
            _lblBasicUser, _basicUser,
            _lblBasicPass, _basicPass,
            _chkShowBasicPass,

            _enabled,
            _lblHotkey, _hotkeyBox,

            _lblLang, _lang,

            _lblCleanup,
            _lblCleanupUrl, _cleanupUrl,
            _lblCleanupToken, _cleanupToken,
            _chkShowCleanupToken,

            _cleanupDailyEnabled, _lblDailyTime, _cleanupDailyHour, _lblH, _cleanupDailyMinute, _lblM,
            _cleanupEveryEnabled, _lblEvery, _cleanupEveryMinutes,

            _btnSave, _btnCancel,

            _bottomRow, // ★最後に追加
        });

        // token表示
        _chkShowToken.CheckedChanged += (_, __) => { _token.UseSystemPasswordChar = !_chkShowToken.Checked; };
        _token.UseSystemPasswordChar = !_chkShowToken.Checked;

        _chkShowCleanupToken.CheckedChanged += (_, __) => { _cleanupToken.UseSystemPasswordChar = !_chkShowCleanupToken.Checked; };
        _cleanupToken.UseSystemPasswordChar = !_chkShowCleanupToken.Checked;

        // BASIC pass 表示
        _chkShowBasicPass.CheckedChanged += (_, __) => { _basicPass.UseSystemPasswordChar = !_chkShowBasicPass.Checked; };
        _basicPass.UseSystemPasswordChar = !_chkShowBasicPass.Checked;

        // ★Helpボタン
        _btnHelp.Click += (_, __) => OpenUrlOrShowError(HelpUrl, SafeT("HelpLink", "Help"));

        // ★Websiteリンク
        _lnkWebsite.LinkClicked += (_, __) => OpenUrlOrShowError(WebsiteUrl, SafeT("WebsiteLink", "GitHub"));

        // ホットキー入力
        _hotkeyBox.KeyDown += HotkeyBox_KeyDown;
        _hotkeyBox.GotFocus += (_, __) => { _hotkeyBox.BackColor = System.Drawing.Color.LightYellow; };
        _hotkeyBox.LostFocus += (_, __) => { _hotkeyBox.BackColor = System.Drawing.SystemColors.Window; };

        // 相互排他 & グレーアウト制御
        _cleanupDailyEnabled.CheckedChanged += (_, __) =>
        {
            if (_cleanupDailyEnabled.Checked) _cleanupEveryEnabled.Checked = false;
            ApplyCleanupUiEnabledState();
        };
        _cleanupEveryEnabled.CheckedChanged += (_, __) =>
        {
            if (_cleanupEveryEnabled.Checked) _cleanupDailyEnabled.Checked = false;
            ApplyCleanupUiEnabledState();
        };

        _btnTest.Click += async (_, __) => await TestConnectionAsync();
        _btnSave.Click += (_, __) => SaveAndClose();
        _btnCancel.Click += (_, __) => Close();

        // 言語リスト
        foreach (var (code, display) in I18n.SupportedLanguages)
        {
            _lang.Items.Add(new LangItem(code, display));
        }
        _lang.SelectedIndexChanged += (_, __) =>
        {
            if (_lang.SelectedItem is LangItem li)
            {
                I18n.SetLanguage(li.Code);
                ApplyLanguageTexts();
            }
        };

        LoadFromSettings();
    }

    private void OpenUrlOrShowError(string url, string title)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void LoadFromSettings()
    {
        var s = SettingsStore.Current;

        var code = string.IsNullOrWhiteSpace(s.Language) ? "en" : s.Language;
        SelectLanguage(code);

        _url.Text = s.BaseUrl ?? "";
        _token.Text = DpapiHelper.Decrypt(s.TokenEncrypted);
        _enabled.Checked = s.Enabled;
        _showSuccess.Checked = s.ShowMessageOnSuccess;

        // BASIC
        _basicUser.Text = s.BasicUser ?? "";
        _basicPass.Text = DpapiHelper.Decrypt(s.BasicPassEncrypted);

        _pendingHotkeyMods = s.HotkeyModifiers;
        _pendingHotkeyVk = s.HotkeyVk;
        _pendingHotkeyDisplay = s.HotkeyDisplay;

        _hotkeyBox.Text = string.IsNullOrWhiteSpace(_pendingHotkeyDisplay)
            ? BuildHotkeyDisplay(_pendingHotkeyMods, (Keys)_pendingHotkeyVk)
            : _pendingHotkeyDisplay;

        _cleanupUrl.Text = s.CleanupBaseUrl ?? "";
        _cleanupToken.Text = DpapiHelper.Decrypt(s.CleanupTokenEncrypted);

        _cleanupDailyEnabled.Checked = s.CleanupDailyEnabled;
        _cleanupDailyHour.Value = s.CleanupDailyHour;
        _cleanupDailyMinute.Value = s.CleanupDailyMinute;

        _cleanupEveryEnabled.Checked = s.CleanupEveryEnabled;
        _cleanupEveryMinutes.Value = s.CleanupEveryMinutes;

        ApplyCleanupUiEnabledState();
        ApplyLanguageTexts();
    }

    private void SelectLanguage(string code)
    {
        for (int i = 0; i < _lang.Items.Count; i++)
        {
            if (_lang.Items[i] is LangItem li && string.Equals(li.Code, code, StringComparison.OrdinalIgnoreCase))
            {
                _lang.SelectedIndex = i;
                return;
            }
        }
        if (_lang.Items.Count > 0) _lang.SelectedIndex = 0;
    }

    private void ApplyLanguageTexts()
    {
        Text = I18n.T("SettingsTitle");

        _lblUrl.Text = I18n.T("PostUrlLabel");
        _lblToken.Text = I18n.T("TokenLabel");
        _chkShowToken.Text = I18n.T("ShowToken");
        _btnTest.Text = I18n.T("TestConnection");

        _showSuccess.Text = I18n.T("ShowSuccessCheckbox");

        // BASIC
        _lblBasic.Text = I18n.T("BasicAuthSection");
        _lblBasicUser.Text = I18n.T("BasicUserLabel");
        _lblBasicPass.Text = I18n.T("BasicPassLabel");
        _chkShowBasicPass.Text = I18n.T("ShowBasicPass");

        _enabled.Text = I18n.T("EnabledCheckbox");
        _lblHotkey.Text = I18n.T("HotkeyLabel");

        _lblLang.Text = I18n.T("LangLabel");

        // ★最下段
        _btnHelp.Text = SafeT("HelpLink", "Help");
        _lblAppName.Text = GetProductName();
        _lblVersion.Text = $"{SafeT("VersionLabel", "Version")}: {GetAppVersion()}";

        // ★追加：Webサイトリンク
        _lnkWebsite.Text = SafeT("WebsiteLink", "GitHub");

        _lblCleanup.Text = I18n.T("CleanupSection");
        _lblCleanupUrl.Text = I18n.T("CleanupUrlLabel");
        _lblCleanupToken.Text = I18n.T("CleanupTokenLabel");
        _chkShowCleanupToken.Text = I18n.T("ShowCleanupToken");

        _cleanupDailyEnabled.Text = I18n.T("DailyCheckbox");
        _lblDailyTime.Text = I18n.T("TimeLabel");
        _lblH.Text = I18n.T("Hour");
        _lblM.Text = I18n.T("Minute");

        _cleanupEveryEnabled.Text = I18n.T("EveryCheckbox");
        _lblEvery.Text = I18n.T("EveryMinutesLabel");

        _btnSave.Text = I18n.T("Save");
        _btnCancel.Text = I18n.T("Cancel");
    }

    private static string SafeT(string key, string fallback)
    {
        var v = I18n.T(key);
        return string.Equals(v, key, StringComparison.OrdinalIgnoreCase) ? fallback : v;
    }

    private static string GetProductName()
    {
        try
        {
            var asm = Assembly.GetExecutingAssembly();

            // AssemblyProductAttribute が入っていればそれ
            var prod = asm.GetCustomAttribute<AssemblyProductAttribute>()?.Product;
            if (!string.IsNullOrWhiteSpace(prod)) return prod.Trim();

            // なければ AssemblyName
            return asm.GetName().Name ?? "App";
        }
        catch
        {
            return "App";
        }
    }

    private static string GetAppVersion()
    {
        try
        {
            var asm = Assembly.GetExecutingAssembly();

            // InformationalVersion（例: 1.2.3 / 1.2.3+abcdef）
            var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            if (!string.IsNullOrWhiteSpace(info)) return info.Trim();

            // AssemblyVersion
            var v = asm.GetName().Version;
            return v?.ToString() ?? "unknown";
        }
        catch
        {
            return "unknown";
        }
    }

    private void HotkeyBox_KeyDown(object? sender, KeyEventArgs e)
    {
        e.SuppressKeyPress = true;

        uint mods = 0;
        if (e.Control) mods |= MOD_CONTROL;
        if (e.Alt) mods |= MOD_ALT;
        if (e.Shift) mods |= MOD_SHIFT;

        if (e.KeyCode is Keys.ControlKey or Keys.Menu or Keys.ShiftKey)
            return;

        if (mods == 0)
        {
            MessageBox.Show(this, I18n.T("NeedModifier"), I18n.T("InputErrorTitle"),
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        int vk = (int)e.KeyCode;
        var display = BuildHotkeyDisplay(mods, e.KeyCode);

        _pendingHotkeyMods = mods;
        _pendingHotkeyVk = vk;
        _pendingHotkeyDisplay = display;
        _hotkeyBox.Text = display;
    }

    private static string BuildHotkeyDisplay(uint mods, Keys key)
    {
        var parts = new List<string>();
        if ((mods & MOD_CONTROL) != 0) parts.Add("Ctrl");
        if ((mods & MOD_ALT) != 0) parts.Add("Alt");
        if ((mods & MOD_SHIFT) != 0) parts.Add("Shift");
        parts.Add(key.ToString());
        return string.Join(" + ", parts);
    }

    // 接続テストもBASIC認証対応
    private async Task TestConnectionAsync()
    {
        var url = (_url.Text ?? "").Trim();
        var token = (_token.Text ?? "").Trim();

        var basicUser = (_basicUser.Text ?? "").Trim();
        var basicPass = (_basicPass.Text ?? "");

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
            (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp))
        {
            MessageBox.Show(this, I18n.T("UrlInvalid"), I18n.T("TestConnection"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (string.IsNullOrWhiteSpace(token))
        {
            MessageBox.Show(this, I18n.T("TokenEmpty"), I18n.T("TestConnection"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (!string.IsNullOrWhiteSpace(basicUser) && string.IsNullOrEmpty(basicPass))
        {
            MessageBox.Show(this, I18n.T("BasicIncomplete"), I18n.T("InputErrorTitle"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var testText = "TEST\n" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };

            var kv = new List<KeyValuePair<string, string>>
            {
                new("token", token),
                new("text", testText),
            };

            using var content = new FormUrlEncodedContent(kv);
            content.Headers.ContentType!.CharSet = "utf-8";

            using var req = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = content
            };

            if (!string.IsNullOrWhiteSpace(basicUser))
            {
                var raw = $"{basicUser}:{basicPass}";
                var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(raw));
                req.Headers.Authorization = new AuthenticationHeaderValue("Basic", b64);
            }

            using var res = await client.SendAsync(req);
            var body = await res.Content.ReadAsStringAsync();

            if (!res.IsSuccessStatusCode)
            {
                MessageBox.Show(this, $"HTTP {(int)res.StatusCode}\n\n{body}", I18n.T("TestConnection"),
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show(this,
                string.IsNullOrWhiteSpace(body) ? "OK" : body,
                I18n.T("TestConnection"),
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, I18n.T("TestConnection"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ApplyCleanupUiEnabledState()
    {
        var daily = _cleanupDailyEnabled.Checked;
        _cleanupDailyHour.Enabled = daily;
        _cleanupDailyMinute.Enabled = daily;

        var every = _cleanupEveryEnabled.Checked;
        _cleanupEveryMinutes.Enabled = every;
    }

    // 保存にBASIC認証を追加
    private void SaveAndClose()
    {
        var url = (_url.Text ?? "").Trim();
        var token = (_token.Text ?? "").Trim();

        var basicUser = (_basicUser.Text ?? "").Trim();
        var basicPass = (_basicPass.Text ?? "");

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
            (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp))
        {
            MessageBox.Show(this, I18n.T("UrlInvalid"),
                I18n.T("InputErrorTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (string.IsNullOrWhiteSpace(token))
        {
            MessageBox.Show(this, I18n.T("TokenEmpty"), I18n.T("InputErrorTitle"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (!string.IsNullOrWhiteSpace(basicUser) && string.IsNullOrEmpty(basicPass))
        {
            MessageBox.Show(this, I18n.T("BasicIncomplete"), I18n.T("InputErrorTitle"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var langCode = (_lang.SelectedItem as LangItem)?.Code ?? "en";

        SettingsStore.Save(new AppSettings
        {
            BaseUrl = url,
            TokenEncrypted = DpapiHelper.Encrypt(token),
            Enabled = _enabled.Checked,
            ShowMessageOnSuccess = _showSuccess.Checked,
            HotkeyModifiers = _pendingHotkeyMods,
            HotkeyVk = _pendingHotkeyVk,
            HotkeyDisplay = _pendingHotkeyDisplay,
            CleanupBaseUrl = (_cleanupUrl.Text ?? "").Trim(),
            CleanupTokenEncrypted = DpapiHelper.Encrypt((_cleanupToken.Text ?? "").Trim()),
            CleanupDailyEnabled = _cleanupDailyEnabled.Checked,
            CleanupDailyHour = (int)_cleanupDailyHour.Value,
            CleanupDailyMinute = (int)_cleanupDailyMinute.Value,
            CleanupEveryEnabled = _cleanupEveryEnabled.Checked,
            CleanupEveryMinutes = (int)_cleanupEveryMinutes.Value,
            Language = langCode,

            // BASIC認証（共通）
            BasicUser = basicUser,
            BasicPassEncrypted = DpapiHelper.Encrypt(basicPass),
        });

        MessageBox.Show(this, I18n.T("SavedMsg"), I18n.T("SavedTitle"),
            MessageBoxButtons.OK, MessageBoxIcon.Information);

        Close();
    }

    private sealed class LangItem
    {
        public string Code { get; }
        public string Display { get; }
        public LangItem(string code, string display) { Code = code; Display = display; }
        public override string ToString() => Display;
    }
}