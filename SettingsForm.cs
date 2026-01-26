// SettingsForm.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;

internal sealed class SettingsForm : Form
{
    private const string HelpUrl = "https://stayhomelab.net/ClipboardSync";
    private const string WebsiteUrl = "https://github.com/StayHomeLabNet/ClipboardSync";

    // MOD_*（Program.csと合わせる）
    private const uint MOD_ALT = 0x0001;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;

    // ======= Tabs =======
    private readonly TabControl _tabs = new()
    {
        Dock = DockStyle.Fill
    };

    private readonly TabPage _tabSend = new();
    private readonly TabPage _tabReceive = new();

    // ======= 共通（tokenは送信/受信で同じ） =======
    private readonly TextBox _tokenSend = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, UseSystemPasswordChar = true };
    private readonly CheckBox _chkShowTokenSend = new() { AutoSize = true, Checked = false };

    private readonly TextBox _tokenRecv = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, UseSystemPasswordChar = true };
    private readonly CheckBox _chkShowTokenRecv = new() { AutoSize = true, Checked = false };

    private bool _syncingToken = false;

    // ======= 送信タブ =======
    private readonly Label _lblUrl = new() { AutoSize = true, Left = 12, Top = 16 };
    private readonly TextBox _url = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };

    private readonly Label _lblTokenSend = new() { AutoSize = true, Left = 12 };
    private readonly CheckBox _enabled = new() { AutoSize = true };
    private readonly CheckBox _showSuccess = new() { AutoSize = true };
    private readonly Button _btnTest = new() { Width = 120, Height = 30 };

    private readonly Label _lblHotkeySend = new() { AutoSize = true, Left = 12 };
    private readonly TextBox _hotkeySendBox = new() { ReadOnly = true, TabStop = true };

    // BASIC（送信タブ右側）
    private readonly Label _lblBasic = new() { AutoSize = true, Left = 12 };
    private readonly Label _lblBasicUser = new() { AutoSize = true, Left = 12 };
    private readonly TextBox _basicUser = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
    private readonly Label _lblBasicPass = new() { AutoSize = true, Left = 12 };
    private readonly TextBox _basicPass = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, UseSystemPasswordChar = true };
    private readonly CheckBox _chkShowBasicPass = new() { AutoSize = true, Checked = false };

    // Cleanup（送信タブ）
    private readonly Label _lblCleanup = new() { AutoSize = true, Left = 12 };
    private readonly Label _lblCleanupUrl = new() { AutoSize = true, Left = 12 };
    private readonly TextBox _cleanupUrl = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
    private readonly Label _lblCleanupToken = new() { AutoSize = true, Left = 12 };
    private readonly TextBox _cleanupToken = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, UseSystemPasswordChar = true };
    private readonly CheckBox _chkShowCleanupToken = new() { AutoSize = true, Checked = false };

    // 自動削除（送信タブ）
    private readonly CheckBox _cleanupDailyEnabled = new() { AutoSize = true };
    private readonly NumericUpDown _cleanupDailyHour = new() { Minimum = 0, Maximum = 23, Width = 60 };
    private readonly NumericUpDown _cleanupDailyMinute = new() { Minimum = 0, Maximum = 59, Width = 60 };
    private readonly Label _lblDailyTime = new() { AutoSize = true, Left = 30 };
    private readonly Label _lblH = new() { AutoSize = true };
    private readonly Label _lblM = new() { AutoSize = true };

    private readonly CheckBox _cleanupEveryEnabled = new() { AutoSize = true };
    private readonly NumericUpDown _cleanupEveryMinutes = new() { Minimum = 1, Maximum = 1440, Width = 80 };
    private readonly Label _lblEvery = new() { AutoSize = true, Left = 30 };

    // ★追加：バックアップ数（右横表示）＆ 一括削除ボタン
    private readonly Label _lblBakCount = new() { AutoSize = true };
    private readonly Button _btnPurgeBak = new() { Width = 200, Height = 30 };

    private readonly System.Windows.Forms.Timer _bakDebounceTimer = new();
    private int _bakQuerying = 0;

    // 言語（送信タブ）
    private readonly Label _lblLang = new() { AutoSize = true, Left = 12 };
    private readonly ComboBox _lang = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };

    // 送信タブ Save/Cancel
    private readonly Button _btnSaveSend = new() { Width = 120, Height = 30 };
    private readonly Button _btnCancelSend = new() { Width = 120, Height = 30 };

    // ======= 受信タブ =======
    private readonly Label _lblReceiveUrl = new() { AutoSize = true, Left = 12, Top = 16 };
    private readonly TextBox _receiveUrl = new() { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };

    private readonly Label _lblTokenRecv = new() { AutoSize = true, Left = 12 };

    private readonly Label _lblHotkeyRecv = new() { AutoSize = true, Left = 12 };
    private readonly TextBox _hotkeyRecvBox = new() { ReadOnly = true, TabStop = true };

    private readonly CheckBox _receiveAutoPaste = new() { AutoSize = true, Left = 12 };
    private readonly Label _lblStableWait = new() { AutoSize = true, Left = 12 };
    private readonly NumericUpDown _stableWaitMs = new() { Minimum = 0, Maximum = 2000, Width = 90, Left = 12 };

    // 受信タブ Save/Cancel
    private readonly Button _btnSaveRecv = new() { Width = 120, Height = 30 };
    private readonly Button _btnCancelRecv = new() { Width = 120, Height = 30 };

    // ======= 最下段（Help/App/Version/Web） =======
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

    // ======= pending hotkeys =======
    private uint _pendingSendMods;
    private int _pendingSendVk;
    private string _pendingSendDisplay = "";

    private uint _pendingRecvMods;
    private int _pendingRecvVk;
    private string _pendingRecvDisplay = "";

    public SettingsForm()
    {
        // “タブ幅に合わせて伸び縮み” を活かすため、サイズ変更可能にする（推奨）
        Width = 720;
        Height = 960;
        MinimumSize = new System.Drawing.Size(680, 880);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = false;

        // Tabs
        Controls.Add(_tabs);

        _tabs.TabPages.Add(_tabSend);
        _tabs.TabPages.Add(_tabReceive);

        _tabSend.Padding = new Padding(8);
        _tabReceive.Padding = new Padding(8);

        // -------------------------
        // Send tab layout（絶対配置）
        // -------------------------
        _url.Left = 12; _url.Top = 38; _url.Width = 660;

        _lblTokenSend.Left = 12; _lblTokenSend.Top = 78;
        _tokenSend.Left = 12; _tokenSend.Top = 98; _tokenSend.Width = 660;
        _chkShowTokenSend.Left = 12; _chkShowTokenSend.Top = _tokenSend.Bottom + 6;

        _btnTest.Left = 552; _btnTest.Top = _tokenSend.Bottom + 2;

        _showSuccess.Left = 12; _showSuccess.Top = _chkShowTokenSend.Bottom + 10;
        _enabled.Left = 12; _enabled.Top = _showSuccess.Bottom + 10;

        _lblHotkeySend.Top = _enabled.Bottom + 16;
        _hotkeySendBox.Left = 12; _hotkeySendBox.Top = _lblHotkeySend.Bottom + 6; _hotkeySendBox.Width = 300;

        // BASIC（右側）
        _lblBasic.Left = 360; _lblBasic.Top = _showSuccess.Top;
        _lblBasicUser.Left = 360; _lblBasicUser.Top = _lblBasic.Bottom + 8;
        _basicUser.Left = 360; _basicUser.Top = _lblBasicUser.Bottom + 6; _basicUser.Width = 312;

        _lblBasicPass.Left = 360; _lblBasicPass.Top = _basicUser.Bottom + 10;
        _basicPass.Left = 360; _basicPass.Top = _lblBasicPass.Bottom + 6; _basicPass.Width = 312;

        _chkShowBasicPass.Left = 360; _chkShowBasicPass.Top = _basicPass.Bottom + 6;

        // Cleanup
        _lblCleanup.Left = 12; _lblCleanup.Top = _hotkeySendBox.Bottom + 18;

        _lblCleanupUrl.Left = 12; _lblCleanupUrl.Top = _lblCleanup.Bottom + 8;
        _cleanupUrl.Left = 12; _cleanupUrl.Top = _lblCleanupUrl.Bottom + 6; _cleanupUrl.Width = 660;

        _lblCleanupToken.Left = 12; _lblCleanupToken.Top = _cleanupUrl.Bottom + 10;
        _cleanupToken.Left = 12; _cleanupToken.Top = _lblCleanupToken.Bottom + 6; _cleanupToken.Width = 660;

        _chkShowCleanupToken.Left = 12; _chkShowCleanupToken.Top = _cleanupToken.Bottom + 6;

        _cleanupDailyEnabled.Left = 12; _cleanupDailyEnabled.Top = _chkShowCleanupToken.Bottom + 12;

        // ★右横にバックアップ数
        _lblBakCount.Left = _lblBasic.Left;
        _lblBakCount.Top = _cleanupDailyEnabled.Top + 2;

        // ★その下に一括削除ボタン
        _btnPurgeBak.Left = _lblBasic.Left;
        _btnPurgeBak.Top = _cleanupDailyEnabled.Bottom + 6;

        // daily time controls はボタンの下へ
        _lblDailyTime.Left = 30; _lblDailyTime.Top = _btnPurgeBak.Bottom + 8;
        _cleanupDailyHour.Left = 30; _cleanupDailyHour.Top = _lblDailyTime.Bottom + 4;
        _lblH.Left = _cleanupDailyHour.Right + 6; _lblH.Top = _cleanupDailyHour.Top + 4;
        _cleanupDailyMinute.Left = _lblH.Right + 10; _cleanupDailyMinute.Top = _cleanupDailyHour.Top;
        _lblM.Left = _cleanupDailyMinute.Right + 6; _lblM.Top = _cleanupDailyMinute.Top + 4;

        _cleanupEveryEnabled.Left = 12; _cleanupEveryEnabled.Top = _cleanupDailyHour.Bottom + 14;
        _lblEvery.Left = 30; _lblEvery.Top = _cleanupEveryEnabled.Bottom + 8;
        _cleanupEveryMinutes.Left = 30; _cleanupEveryMinutes.Top = _lblEvery.Bottom + 4;

        // Language
        _lblLang.Left = 12; _lblLang.Top = _cleanupEveryMinutes.Bottom + 18;
        _lang.Left = 12; _lang.Top = _lblLang.Bottom + 6;

        // Save/Cancel (send tab)
        _btnSaveSend.Left = 552; _btnSaveSend.Top = _lang.Top;
        _btnCancelSend.Left = 424; _btnCancelSend.Top = _lang.Top;

        // -------------------------
        // Receive tab layout
        // -------------------------
        _receiveUrl.Left = 12; _receiveUrl.Top = 38; _receiveUrl.Width = 660;

        _lblTokenRecv.Left = 12; _lblTokenRecv.Top = 78;
        _tokenRecv.Left = 12; _tokenRecv.Top = 98; _tokenRecv.Width = 660;
        _chkShowTokenRecv.Left = 12; _chkShowTokenRecv.Top = _tokenRecv.Bottom + 6;

        _lblHotkeyRecv.Left = 12; _lblHotkeyRecv.Top = _chkShowTokenRecv.Bottom + 14;
        _hotkeyRecvBox.Left = 12; _hotkeyRecvBox.Top = _lblHotkeyRecv.Bottom + 6; _hotkeyRecvBox.Width = 300;

        _receiveAutoPaste.Left = 12;
        _receiveAutoPaste.Top = _hotkeyRecvBox.Bottom + 16;

        _lblStableWait.Left = 12;
        _lblStableWait.Top = _receiveAutoPaste.Bottom + 12;

        _stableWaitMs.Left = 12;
        _stableWaitMs.Top = _lblStableWait.Bottom + 6;

        // Receive tab Save/Cancel
        _btnSaveRecv.Left = 552;
        _btnCancelRecv.Left = 424;
        _btnSaveRecv.Top = _stableWaitMs.Bottom + 18;
        _btnCancelRecv.Top = _btnSaveRecv.Top;

        // -------------------------
        // Bottom row
        // -------------------------
        _bottomRow.Controls.Add(_btnHelp);

        _lblAppName.Margin = new Padding(14, 6, 0, 0);
        _lblVersion.Margin = new Padding(10, 6, 0, 0);
        _lnkWebsite.Margin = new Padding(10, 6, 0, 0);

        _bottomRow.Controls.Add(_lblAppName);
        _bottomRow.Controls.Add(_lblVersion);
        _bottomRow.Controls.Add(_lnkWebsite);

        Controls.Add(_bottomRow);

        // Add controls into tabs
        _tabSend.Controls.AddRange(new Control[]
        {
            _lblUrl, _url,

            _lblTokenSend, _tokenSend, _chkShowTokenSend,
            _btnTest,
            _showSuccess,
            _enabled,

            _lblHotkeySend, _hotkeySendBox,

            _lblBasic, _lblBasicUser, _basicUser, _lblBasicPass, _basicPass, _chkShowBasicPass,

            _lblCleanup, _lblCleanupUrl, _cleanupUrl, _lblCleanupToken, _cleanupToken, _chkShowCleanupToken,

            _cleanupDailyEnabled, _lblBakCount, _btnPurgeBak, // ★追加

            _lblDailyTime, _cleanupDailyHour, _lblH, _cleanupDailyMinute, _lblM,

            _cleanupEveryEnabled, _lblEvery, _cleanupEveryMinutes,

            _lblLang, _lang,
            _btnSaveSend, _btnCancelSend,
        });

        _tabReceive.Controls.AddRange(new Control[]
        {
            _lblReceiveUrl, _receiveUrl,

            _lblTokenRecv, _tokenRecv, _chkShowTokenRecv,

            _lblHotkeyRecv, _hotkeyRecvBox,

            _receiveAutoPaste,
            _lblStableWait, _stableWaitMs,

            _btnSaveRecv, _btnCancelRecv,
        });

        // Events: show/hide token
        _chkShowTokenSend.CheckedChanged += (_, __) => _tokenSend.UseSystemPasswordChar = !_chkShowTokenSend.Checked;
        _chkShowTokenRecv.CheckedChanged += (_, __) => _tokenRecv.UseSystemPasswordChar = !_chkShowTokenRecv.Checked;

        _chkShowCleanupToken.CheckedChanged += (_, __) => _cleanupToken.UseSystemPasswordChar = !_chkShowCleanupToken.Checked;
        _chkShowBasicPass.CheckedChanged += (_, __) => _basicPass.UseSystemPasswordChar = !_chkShowBasicPass.Checked;

        // Token sync（送信と受信の token は同じなので同期）
        _tokenSend.TextChanged += (_, __) =>
        {
            if (_syncingToken) return;
            _syncingToken = true;
            _tokenRecv.Text = _tokenSend.Text;
            _syncingToken = false;
        };
        _tokenRecv.TextChanged += (_, __) =>
        {
            if (_syncingToken) return;
            _syncingToken = true;
            _tokenSend.Text = _tokenRecv.Text;
            _syncingToken = false;
        };

        // Help & Website
        _btnHelp.Click += (_, __) => OpenUrlOrShowError(HelpUrl, SafeT("HelpLink", "Help"));
        _lnkWebsite.LinkClicked += (_, __) => OpenUrlOrShowError(WebsiteUrl, SafeT("WebsiteLink", "GitHub"));

        // Hotkey boxes
        _hotkeySendBox.KeyDown += (_, e) => HotkeyBox_KeyDown_Send(e);
        _hotkeySendBox.GotFocus += (_, __) => { _hotkeySendBox.BackColor = System.Drawing.Color.LightYellow; };
        _hotkeySendBox.LostFocus += (_, __) => { _hotkeySendBox.BackColor = System.Drawing.SystemColors.Window; };

        _hotkeyRecvBox.KeyDown += (_, e) => HotkeyBox_KeyDown_Recv(e);
        _hotkeyRecvBox.GotFocus += (_, __) => { _hotkeyRecvBox.BackColor = System.Drawing.Color.LightYellow; };
        _hotkeyRecvBox.LostFocus += (_, __) => { _hotkeyRecvBox.BackColor = System.Drawing.SystemColors.Window; };

        // Mutual exclusion for cleanup schedules
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

        // Buttons
        _btnTest.Click += async (_, __) => await TestConnectionAsync();
        _btnPurgeBak.Click += async (_, __) => await PurgeBackupsAsync();

        _btnSaveSend.Click += (_, __) => SaveAndClose();
        _btnCancelSend.Click += (_, __) => Close();

        _btnSaveRecv.Click += (_, __) => SaveAndClose();
        _btnCancelRecv.Click += (_, __) => Close();

        // Lang list
        foreach (var (code, display) in I18n.SupportedLanguages)
            _lang.Items.Add(new LangItem(code, display));

        // ★言語変更は「即反映」＋ UI再レイアウト
        _lang.SelectedIndexChanged += (_, __) =>
        {
            if (_lang.SelectedItem is LangItem li)
            {
                I18n.SetLanguage(li.Code);
                ApplyLanguageTexts();
                ApplyResponsiveLayout();
                _ = RefreshBackupCountAsync(); // 言語表示も変わるので更新
            }
        };

        // ★バックアップ数 取得：URL/Token変更で取り直し（デバウンス）
        _bakDebounceTimer.Interval = 650;
        _bakDebounceTimer.Tick += async (_, __) =>
        {
            _bakDebounceTimer.Stop();
            await RefreshBackupCountAsync();
        };

        void scheduleBakRefresh()
        {
            _bakDebounceTimer.Stop();
            _bakDebounceTimer.Start();
        }

        _cleanupUrl.TextChanged += (_, __) => scheduleBakRefresh();
        _cleanupToken.TextChanged += (_, __) => scheduleBakRefresh();
        _chkShowCleanupToken.CheckedChanged += (_, __) => scheduleBakRefresh();

        // ★タブ幅に追従
        Shown += async (_, __) =>
        {
            ApplyResponsiveLayout();
            await RefreshBackupCountAsync();
        };

        _tabSend.Resize += (_, __) => ApplyResponsiveLayout();
        _tabReceive.Resize += (_, __) => ApplyResponsiveLayout();
        Resize += (_, __) => ApplyResponsiveLayout();

        LoadFromSettings();
        ApplyResponsiveLayout();
    }

    private void ApplyResponsiveLayout()
    {
        // タブの“内側”の幅で合わせる（Paddingの影響を吸収）
        FitWidthToTab(_tabSend, _url, 12, 12);
        FitWidthToTab(_tabSend, _tokenSend, 12, 12);
        FitWidthToTab(_tabSend, _cleanupUrl, 12, 12);
        FitWidthToTab(_tabSend, _cleanupToken, 12, 12);

        FitWidthToTab(_tabReceive, _receiveUrl, 12, 12);
        FitWidthToTab(_tabReceive, _tokenRecv, 12, 12);

        // BASICは右側配置なので、フォーム幅に応じて右の幅を可変にする（余白を維持）
        var tabW = _tabSend.ClientSize.Width;
        var rightLeft = 360;
        var rightMargin = 12;
        var rightW = Math.Max(120, tabW - rightLeft - rightMargin);
        _basicUser.Width = rightW;
        _basicPass.Width = rightW;

        // テストボタンは右端へ
        _btnTest.Left = Math.Max(12, _tabSend.ClientSize.Width - 12 - _btnTest.Width);

        // ★バックアップ数ラベル：右端に寄せて「右横」感を出す
        var xRight = _tabSend.ClientSize.Width - 12 - _lblBakCount.Width;
        var xMin = _cleanupDailyEnabled.Right + 12;
        _lblBakCount.Left = Math.Max(xMin, xRight);
        _lblBakCount.Top = _cleanupDailyEnabled.Top + 2;

        // purgeボタンの幅も伸縮（任意）
        _btnPurgeBak.Width = Math.Min(260, Math.Max(160, _tabSend.ClientSize.Width - 24));

        // -------------------------
        // 位置調整（ここに置くと常に維持される）
        // -------------------------
        _lblBakCount.Left = _lblBasic.Left;
        _lblDailyTime.Top = _btnPurgeBak.Top;
        _cleanupDailyHour.Top = _btnPurgeBak.Bottom + 4;
        _lblH.Top = _cleanupDailyHour.Top + 4;
        _cleanupDailyMinute.Top = _cleanupDailyHour.Top;
        _lblM.Top = _cleanupDailyMinute.Top + 4;
        _cleanupEveryEnabled.Top = _cleanupDailyHour.Bottom + 14;
        _lblEvery.Top = _cleanupEveryEnabled.Bottom + 8;
        _cleanupEveryMinutes.Top = _lblEvery.Bottom + 4;
        _lblLang.Top = _cleanupEveryMinutes.Bottom + 18;
        _lang.Top = _lblLang.Bottom + 6;
        _btnSaveSend.Top = _lang.Top;
        _btnCancelSend.Top = _lang.Top;
    }

    private static void FitWidthToTab(TabPage tab, Control c, int marginLeft, int marginRight)
    {
        var w = tab.ClientSize.Width - marginLeft - marginRight;
        if (w < 50) w = 50;
        c.Left = marginLeft;
        c.Width = w;
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

        // Language first
        var code = string.IsNullOrWhiteSpace(s.Language) ? "en" : s.Language;
        SelectLanguage(code);

        // Send
        _url.Text = s.BaseUrl ?? "";
        var tokenPlain = DpapiHelper.Decrypt(s.TokenEncrypted) ?? "";
        _tokenSend.Text = tokenPlain;

        _enabled.Checked = s.Enabled;
        _showSuccess.Checked = s.ShowMessageOnSuccess;

        _pendingSendMods = s.HotkeyModifiers;
        _pendingSendVk = s.HotkeyVk;
        _pendingSendDisplay = s.HotkeyDisplay;

        _hotkeySendBox.Text = string.IsNullOrWhiteSpace(_pendingSendDisplay)
            ? BuildHotkeyDisplay(_pendingSendMods, (Keys)_pendingSendVk)
            : _pendingSendDisplay;

        // Receive
        _receiveUrl.Text = s.ReceiveBaseUrl ?? "";
        _tokenRecv.Text = tokenPlain; // 同じtoken

        _pendingRecvMods = s.ReceiveHotkeyModifiers;
        _pendingRecvVk = s.ReceiveHotkeyVk;
        _pendingRecvDisplay = s.ReceiveHotkeyDisplay;

        _hotkeyRecvBox.Text = string.IsNullOrWhiteSpace(_pendingRecvDisplay)
            ? BuildHotkeyDisplay(_pendingRecvMods, (Keys)_pendingRecvVk)
            : _pendingRecvDisplay;

        // 追加設定
        _receiveAutoPaste.Checked = s.ReceiveAutoPaste;
        var wait = s.ClipboardStableWaitMs;
        if (wait < (int)_stableWaitMs.Minimum) wait = (int)_stableWaitMs.Minimum;
        if (wait > (int)_stableWaitMs.Maximum) wait = (int)_stableWaitMs.Maximum;
        _stableWaitMs.Value = wait;

        // Basic
        _basicUser.Text = s.BasicUser ?? "";
        _basicPass.Text = DpapiHelper.Decrypt(s.BasicPassEncrypted) ?? "";

        // Cleanup
        _cleanupUrl.Text = s.CleanupBaseUrl ?? "";
        _cleanupToken.Text = DpapiHelper.Decrypt(s.CleanupTokenEncrypted) ?? "";

        _cleanupDailyEnabled.Checked = s.CleanupDailyEnabled;
        _cleanupDailyHour.Value = s.CleanupDailyHour;
        _cleanupDailyMinute.Value = s.CleanupDailyMinute;

        _cleanupEveryEnabled.Checked = s.CleanupEveryEnabled;
        _cleanupEveryMinutes.Value = s.CleanupEveryMinutes;

        ApplyCleanupUiEnabledState();
        ApplyLanguageTexts();

        // 初期表示
        _lblBakCount.Text = I18n.T("BakCountNone");
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

        _tabSend.Text = I18n.T("TabSendDelete");
        _tabReceive.Text = I18n.T("TabReceive");

        // Send tab texts
        _lblUrl.Text = I18n.T("PostUrlLabel");
        _lblTokenSend.Text = I18n.T("TokenLabel");
        _chkShowTokenSend.Text = I18n.T("ShowToken");
        _btnTest.Text = I18n.T("TestConnection");

        _enabled.Text = I18n.T("EnabledCheckbox");
        _showSuccess.Text = I18n.T("ShowSuccessCheckbox");
        _lblHotkeySend.Text = I18n.T("HotkeyLabel");

        _lblBasic.Text = I18n.T("BasicAuthSection");
        _lblBasicUser.Text = I18n.T("BasicUserLabel");
        _lblBasicPass.Text = I18n.T("BasicPassLabel");
        _chkShowBasicPass.Text = I18n.T("ShowBasicPass");

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

        _btnPurgeBak.Text = I18n.T("PurgeBakButton");

        _lblLang.Text = I18n.T("LangLabel");

        _btnSaveSend.Text = I18n.T("Save");
        _btnCancelSend.Text = I18n.T("Cancel");

        // Receive tab texts
        _lblReceiveUrl.Text = I18n.T("ReceiveUrlLabel");
        _lblTokenRecv.Text = I18n.T("TokenLabel");
        _chkShowTokenRecv.Text = I18n.T("ShowToken");
        _lblHotkeyRecv.Text = I18n.T("ReceiveHotkeyLabel");

        _receiveAutoPaste.Text = SafeT("ReceiveAutoPaste", "Auto paste (Ctrl+V) after receive");
        _lblStableWait.Text = SafeT("ClipboardStableWaitLabel", "Clipboard stable wait (ms)");

        _btnSaveRecv.Text = I18n.T("Save");
        _btnCancelRecv.Text = I18n.T("Cancel");

        // Bottom row
        _btnHelp.Text = SafeT("HelpLink", "Help");
        _lblAppName.Text = GetProductName();
        _lblVersion.Text = $"{SafeT("VersionLabel", "Version")}: {GetAppVersion()}";
        _lnkWebsite.Text = SafeT("WebsiteLink", "GitHub");
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
            var prod = asm.GetCustomAttribute<AssemblyProductAttribute>()?.Product;
            if (!string.IsNullOrWhiteSpace(prod)) return prod.Trim();
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
            var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            if (!string.IsNullOrWhiteSpace(info)) return info.Trim();

            var v = asm.GetName().Version;
            return v?.ToString() ?? "unknown";
        }
        catch
        {
            return "unknown";
        }
    }

    private void HotkeyBox_KeyDown_Send(KeyEventArgs e)
    {
        if (!TryCaptureHotkey(e, out var mods, out var vk, out var display)) return;

        _pendingSendMods = mods;
        _pendingSendVk = vk;
        _pendingSendDisplay = display;
        _hotkeySendBox.Text = display;
    }

    private void HotkeyBox_KeyDown_Recv(KeyEventArgs e)
    {
        if (!TryCaptureHotkey(e, out var mods, out var vk, out var display)) return;

        _pendingRecvMods = mods;
        _pendingRecvVk = vk;
        _pendingRecvDisplay = display;
        _hotkeyRecvBox.Text = display;
    }

    private bool TryCaptureHotkey(KeyEventArgs e, out uint mods, out int vk, out string display)
    {
        e.SuppressKeyPress = true;

        mods = 0;
        if (e.Control) mods |= MOD_CONTROL;
        if (e.Alt) mods |= MOD_ALT;
        if (e.Shift) mods |= MOD_SHIFT;

        display = "";
        vk = 0;

        if (e.KeyCode is Keys.ControlKey or Keys.Menu or Keys.ShiftKey)
            return false;

        if (mods == 0)
        {
            MessageBox.Show(this, I18n.T("NeedModifier"), I18n.T("InputErrorTitle"),
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        vk = (int)e.KeyCode;
        display = BuildHotkeyDisplay(mods, e.KeyCode);
        return true;
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

    // Send test（BASIC対応）
    private async Task TestConnectionAsync()
    {
        var url = (_url.Text ?? "").Trim();
        var token = (_tokenSend.Text ?? "").Trim();

        var basicUser = (_basicUser.Text ?? "").Trim();
        var basicPass = (_basicPass.Text ?? "");

        if (string.IsNullOrWhiteSpace(url))
        {
            MessageBox.Show(this, I18n.T("UrlInvalid"), I18n.T("TestConnection"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

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

            using var req = new HttpRequestMessage(HttpMethod.Post, url) { Content = content };

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

    // ★バックアップ数取得
    private async Task RefreshBackupCountAsync()
    {
        if (Interlocked.Exchange(ref _bakQuerying, 1) == 1) return;

        try
        {
            // 入力が空のときは「-」
            if (string.IsNullOrWhiteSpace(_cleanupUrl.Text) || string.IsNullOrWhiteSpace(_cleanupToken.Text))
            {
                _lblBakCount.Text = I18n.T("BakCountNone");
                ApplyResponsiveLayout();
                return;
            }

            _lblBakCount.Text = I18n.T("BakCountLoading");
            ApplyResponsiveLayout();

            var (ok, count, info) = await CleanupApi.GetBackupCountAsync();
            if (ok && count >= 0)
            {
                _lblBakCount.Text = string.Format(I18n.T("BakCountFormat"), count);
            }
            else
            {
                _lblBakCount.Text = I18n.T("BakCountFail");
            }

            ApplyResponsiveLayout();
        }
        finally
        {
            Interlocked.Exchange(ref _bakQuerying, 0);
        }
    }

    // ★バックアップ一括削除
    private async Task PurgeBackupsAsync()
    {
        var r = MessageBox.Show(
            I18n.T("ConfirmPurgeBakBody"),
            I18n.T("ConfirmTitle"),
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);

        if (r != DialogResult.Yes) return;

        _btnPurgeBak.Enabled = false;
        try
        {
            var (ok, info) = await CleanupApi.PurgeBackupsAsync();

            using var top = new Form { TopMost = true, ShowInTaskbar = false };
            top.StartPosition = FormStartPosition.Manual;
            top.Location = new System.Drawing.Point(-2000, -2000);
            top.Show();

            MessageBox.Show(top, info,
                ok ? I18n.T("PurgeBakDoneTitle") : I18n.T("PurgeBakFailTitle"),
                MessageBoxButtons.OK,
                ok ? MessageBoxIcon.Information : MessageBoxIcon.Error);

            await RefreshBackupCountAsync();
        }
        finally
        {
            _btnPurgeBak.Enabled = true;
        }
    }

    // 保存（送信のみ/受信のみでも保存できるように、URLが空なら検証しない）
    private void SaveAndClose()
    {
        var urlSend = (_url.Text ?? "").Trim();
        var urlRecv = (_receiveUrl.Text ?? "").Trim();
        var token = (_tokenSend.Text ?? "").Trim(); // 両方同期されている前提

        var cleanupUrl = (_cleanupUrl.Text ?? "").Trim();
        var cleanupToken = (_cleanupToken.Text ?? "").Trim();

        var basicUser = (_basicUser.Text ?? "").Trim();
        var basicPass = (_basicPass.Text ?? "");

        // URL validation（空ならOK）
        if (!string.IsNullOrWhiteSpace(urlSend))
        {
            if (!Uri.TryCreate(urlSend, UriKind.Absolute, out var u) ||
                (u.Scheme != Uri.UriSchemeHttps && u.Scheme != Uri.UriSchemeHttp))
            {
                MessageBox.Show(this, I18n.T("UrlInvalid"),
                    I18n.T("InputErrorTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        if (!string.IsNullOrWhiteSpace(urlRecv))
        {
            if (!Uri.TryCreate(urlRecv, UriKind.Absolute, out var u) ||
                (u.Scheme != Uri.UriSchemeHttps && u.Scheme != Uri.UriSchemeHttp))
            {
                MessageBox.Show(this, I18n.T("UrlInvalid"),
                    I18n.T("InputErrorTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        // token required if either URL is set
        if ((!string.IsNullOrWhiteSpace(urlSend) || !string.IsNullOrWhiteSpace(urlRecv)) &&
            string.IsNullOrWhiteSpace(token))
        {
            MessageBox.Show(this, I18n.T("TokenEmpty"), I18n.T("InputErrorTitle"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // cleanup optional, but if set then validate token too
        if (!string.IsNullOrWhiteSpace(cleanupUrl))
        {
            if (!Uri.TryCreate(cleanupUrl, UriKind.Absolute, out var u) ||
                (u.Scheme != Uri.UriSchemeHttps && u.Scheme != Uri.UriSchemeHttp))
            {
                MessageBox.Show(this, I18n.T("UrlInvalid"),
                    I18n.T("InputErrorTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(cleanupToken))
            {
                MessageBox.Show(this, I18n.T("TokenEmpty"), I18n.T("InputErrorTitle"),
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        // BASIC validation
        if (!string.IsNullOrWhiteSpace(basicUser) && string.IsNullOrEmpty(basicPass))
        {
            MessageBox.Show(this, I18n.T("BasicIncomplete"), I18n.T("InputErrorTitle"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var langCode = (_lang.SelectedItem as LangItem)?.Code ?? "en";

        SettingsStore.Save(new AppSettings
        {
            // Send
            BaseUrl = urlSend,
            TokenEncrypted = DpapiHelper.Encrypt(token),
            Enabled = _enabled.Checked,
            ShowMessageOnSuccess = _showSuccess.Checked,
            HotkeyModifiers = _pendingSendMods,
            HotkeyVk = _pendingSendVk,
            HotkeyDisplay = _pendingSendDisplay,

            // Cleanup
            CleanupBaseUrl = cleanupUrl,
            CleanupTokenEncrypted = DpapiHelper.Encrypt(cleanupToken),
            CleanupDailyEnabled = _cleanupDailyEnabled.Checked,
            CleanupDailyHour = (int)_cleanupDailyHour.Value,
            CleanupDailyMinute = (int)_cleanupDailyMinute.Value,
            CleanupEveryEnabled = _cleanupEveryEnabled.Checked,
            CleanupEveryMinutes = (int)_cleanupEveryMinutes.Value,

            // Receive
            ReceiveBaseUrl = urlRecv,
            ReceiveHotkeyModifiers = _pendingRecvMods,
            ReceiveHotkeyVk = _pendingRecvVk,
            ReceiveHotkeyDisplay = _pendingRecvDisplay,
            ReceiveAutoPaste = _receiveAutoPaste.Checked,
            ClipboardStableWaitMs = (int)_stableWaitMs.Value,

            // UI
            Language = langCode,

            // Basic
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