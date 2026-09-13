using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using ExcelValidationAddIn.Models;
using ExcelValidationAddIn.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelValidationAddIn.UI
{
    public class ValidationPaneControl : UserControl
    {
        private readonly Excel.Application _app;
        private ValidationQueue _queue;
        private string _lastWorksheetName = string.Empty;
        private bool _sheetChangeHooked;
        private CancellationTokenSource _loginCts;

        // ---- status strip ----
        private readonly Panel _statusPanel;
        private readonly Label _statusLabel;

        // ---- auth section ----
        private readonly Panel _authPanel;
        private readonly Panel _authIdlePanel;
        private readonly Button _loginButton;
        private readonly Label _authErrorLabel;
        private readonly Panel _authWaitingPanel;
        private readonly Label _authWaitingLabel;
        private readonly Button _authCancelButton;

        // ---- tenant section ----
        private readonly Panel _tenantPanel;
        private readonly ComboBox _tenantCombo;
        private readonly Button _confirmTenantButton;
        private readonly Label _tenantErrorLabel;

        // ---- working section ----
        private readonly Panel _workingPanel;
        private readonly Label _activeTenantLabel;
        private readonly LinkLabel _switchTenantLink;
        private readonly LinkLabel _logoutLink;
        private readonly Button _validateNowButton;
        private readonly Label _pendingLabel;
        private readonly Label _emptyLabel;
        private readonly ListView _resultsList;

        private static readonly Color AccentColor = Color.FromArgb(0x1F, 0x5C, 0x4A);
        private static readonly Color OkColor = Color.FromArgb(0x1F, 0x5C, 0x4A);
        private static readonly Color WarnColor = Color.FromArgb(0x8A, 0x5A, 0x00);
        private static readonly Color ErrorColor = Color.FromArgb(0xB3, 0x26, 0x1E);
        private static readonly Color ErrorRowBg = Color.FromArgb(0xFB, 0xEA, 0xEA);
        private static readonly Color WarnRowBg = Color.FromArgb(0xFF, 0xF6, 0xE0);

        public ValidationPaneControl(Excel.Application app)
        {
            _app = app;
            RightToLeft = RightToLeft.Yes;
            Dock = DockStyle.Fill;
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.White;

            // ----- status strip -----
            _statusPanel = new Panel { Dock = DockStyle.Top, Height = 34, BackColor = Color.FromArgb(0xE4, 0xEE, 0xEA) };
            _statusLabel = new Label { Text = "آماده", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 10, 0), ForeColor = AccentColor };
            _statusPanel.Controls.Add(_statusLabel);

            // ----- auth section -----
            _authPanel = new Panel { Dock = DockStyle.Fill, Visible = false, Padding = new Padding(16) };

            _authIdlePanel = new Panel { Dock = DockStyle.Fill, Visible = true };
            var authLead = new Label { Text = "برای اعتبارسنجی خودکار سطرها، ابتدا وارد سیستم حسابداری شوید. ورود در مرورگر پیش‌فرض شما انجام می‌شود.", AutoSize = false, Size = new Size(280, 56), Location = new Point(0, 0) };
            _loginButton = new Button { Text = "ورود به سیستم", Location = new Point(0, 64), Width = 280, Height = 32, BackColor = AccentColor, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            _loginButton.FlatAppearance.BorderSize = 0;
            _loginButton.Click += (s, e) => OnLoginClicked();
            _authErrorLabel = new Label { Location = new Point(0, 104), Size = new Size(280, 40), ForeColor = ErrorColor };
            _authIdlePanel.Controls.AddRange(new Control[] { authLead, _loginButton, _authErrorLabel });

            _authWaitingPanel = new Panel { Dock = DockStyle.Fill, Visible = false };
            _authWaitingLabel = new Label { Text = "در حال انتظار برای ورود در مرورگر...\nپس از ورود موفق، این صفحه به‌صورت خودکار به‌روز می‌شود.", AutoSize = false, Size = new Size(280, 60), Location = new Point(0, 0) };
            _authCancelButton = new Button { Text = "لغو", Location = new Point(0, 68), Width = 140, Height = 30, BackColor = Color.White, ForeColor = AccentColor, FlatStyle = FlatStyle.Flat };
            _authCancelButton.FlatAppearance.BorderColor = AccentColor;
            _authCancelButton.Click += (s, e) => OnCancelLoginClicked();
            _authWaitingPanel.Controls.AddRange(new Control[] { _authWaitingLabel, _authCancelButton });

            _authPanel.Controls.Add(_authWaitingPanel);
            _authPanel.Controls.Add(_authIdlePanel);

            // ----- tenant section -----
            _tenantPanel = new Panel { Dock = DockStyle.Fill, Visible = false, Padding = new Padding(16) };
            var tenantLead = new Label { Text = "شرکت (tenant) مورد نظر برای کار را انتخاب کنید.", AutoSize = false, Size = new Size(280, 30), Location = new Point(16, 16) };
            _tenantCombo = new ComboBox { Location = new Point(16, 50), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList };
            _confirmTenantButton = new Button { Text = "تایید و شروع", Location = new Point(16, 84), Width = 280, Height = 32, BackColor = AccentColor, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            _confirmTenantButton.FlatAppearance.BorderSize = 0;
            _confirmTenantButton.Click += (s, e) => OnConfirmTenantClicked();
            _tenantErrorLabel = new Label { Location = new Point(16, 124), Size = new Size(280, 34), ForeColor = ErrorColor };
            _tenantPanel.Controls.AddRange(new Control[] { tenantLead, _tenantCombo, _confirmTenantButton, _tenantErrorLabel });

            // ----- working section -----
            _workingPanel = new Panel { Dock = DockStyle.Fill, Visible = false };

            var sessionRow = new Panel { Dock = DockStyle.Top, Height = 56, Padding = new Padding(12), BorderStyle = BorderStyle.FixedSingle };
            _activeTenantLabel = new Label { Text = "—", AutoSize = true, Location = new Point(12, 8), Font = new Font(Font, FontStyle.Bold) };
            _switchTenantLink = new LinkLabel { Text = "تغییر شرکت", AutoSize = true, Location = new Point(12, 30), LinkColor = AccentColor };
            _switchTenantLink.LinkClicked += (s, e) => OnSwitchTenantClicked();
            _logoutLink = new LinkLabel { Text = "خروج", AutoSize = true, Location = new Point(110, 30), LinkColor = AccentColor };
            _logoutLink.LinkClicked += (s, e) => OnLogoutClicked();
            sessionRow.Controls.AddRange(new Control[] { _activeTenantLabel, _switchTenantLink, _logoutLink });

            var toolbar = new Panel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(12, 6, 12, 6) };
            _validateNowButton = new Button { Text = "اعتبارسنجی الان", Width = 130, Height = 28, Location = new Point(0, 4), BackColor = Color.White, ForeColor = AccentColor, FlatStyle = FlatStyle.Flat };
            _validateNowButton.FlatAppearance.BorderColor = AccentColor;
            _validateNowButton.Click += (s, e) => _queue?.TriggerNow();
            _pendingLabel = new Label { AutoSize = true, Location = new Point(140, 8), ForeColor = Color.Gray };
            toolbar.Controls.AddRange(new Control[] { _validateNowButton, _pendingLabel });

            _emptyLabel = new Label
            {
                Text = "هیچ خطایی ثبت نشده است. با ویرایش شیت، اعتبارسنجی به‌صورت خودکار انجام می‌شود.",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.Gray,
                Padding = new Padding(20)
            };

            _resultsList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                Visible = false,
                RightToLeftLayout = true
            };
            _resultsList.Columns.Add("ردیف", 50);
            _resultsList.Columns.Add("فیلد", 90);
            _resultsList.Columns.Add("پیام", 220);
            _resultsList.MouseDoubleClick += (s, e) => OnResultActivated();
            _resultsList.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OnResultActivated(); };

            _workingPanel.Controls.Add(_resultsList);
            _workingPanel.Controls.Add(_emptyLabel);
            _workingPanel.Controls.Add(toolbar);
            _workingPanel.Controls.Add(sessionRow);

            Controls.Add(_authPanel);
            Controls.Add(_tenantPanel);
            Controls.Add(_workingPanel);
            Controls.Add(_statusPanel);

            Load += (s, e) => Bootstrap();
            Disposed += (s, e) => { _loginCts?.Cancel(); Teardown(); };
        }

        // ---------- bootstrap / section switching ----------

        private void Bootstrap()
        {
            Session session = AuthService.GetSession();

            if (session == null || !session.IsValid())
            {
                ShowAuthSection();
                return;
            }

            if (!session.HasTenant())
            {
                _ = ShowTenantSectionAsync(session);
                return;
            }

            EnterWorkingState(session);
        }

        private void HideAllSections()
        {
            _authPanel.Visible = false;
            _tenantPanel.Visible = false;
            _workingPanel.Visible = false;
        }

        private void ShowAuthSection()
        {
            HideAllSections();
            _authErrorLabel.Text = string.Empty;
            _authWaitingPanel.Visible = false;
            _authIdlePanel.Visible = true;
            _authPanel.Visible = true;
        }

        private async void OnLoginClicked()
        {
            _authErrorLabel.Text = string.Empty;
            _authIdlePanel.Visible = false;
            _authWaitingPanel.Visible = true;

            _loginCts?.Dispose();
            _loginCts = new CancellationTokenSource();

            try
            {
                Session session = await AuthService.LoginAsync(_loginCts.Token);

                if (InvokeRequired)
                {
                    BeginInvoke(new Action(() => _ = ShowTenantSectionAsync(session)));
                    return;
                }

                await ShowTenantSectionAsync(session);
            }
            catch (OperationCanceledException)
            {
                if (InvokeRequired)
                {
                    BeginInvoke(new Action(OnLoginCanceled));
                    return;
                }

                OnLoginCanceled();
            }
            catch (ApiException apiEx)
            {
                if (InvokeRequired)
                {
                    BeginInvoke(new Action(() => OnLoginFailed(apiEx.Message)));
                    return;
                }

                OnLoginFailed(apiEx.Message);
            }
            catch (Exception)
            {
                if (InvokeRequired)
                {
                    BeginInvoke(new Action(() => OnLoginFailed("ورود ناموفق بود.")));
                    return;
                }

                OnLoginFailed("ورود ناموفق بود.");
            }
        }

        private void OnLoginCanceled()
        {
            _authWaitingPanel.Visible = false;
            _authIdlePanel.Visible = true;
        }

        private void OnLoginFailed(string message)
        {
            _authWaitingPanel.Visible = false;
            _authIdlePanel.Visible = true;
            _authErrorLabel.Text = message;
        }

        private void OnCancelLoginClicked()
        {
            _loginCts?.Cancel();
        }

        private async System.Threading.Tasks.Task ShowTenantSectionAsync(Session session)
        {
            HideAllSections();
            _tenantPanel.Visible = true;
            _tenantErrorLabel.Text = string.Empty;
            _tenantCombo.Items.Clear();
            _tenantCombo.Items.Add("در حال بارگذاری...");
            _tenantCombo.SelectedIndex = 0;
            _tenantCombo.Enabled = false;

            List<Tenant> tenants;
            try
            {
                tenants = await AuthService.FetchTenantsAsync(session.Token).ConfigureAwait(true);
            }
            catch (ApiException apiEx)
            {
                _tenantErrorLabel.Text = apiEx.Message;
                _tenantCombo.Items.Clear();
                return;
            }
            catch (Exception)
            {
                _tenantErrorLabel.Text = "بارگذاری لیست شرکت‌ها ناموفق بود.";
                _tenantCombo.Items.Clear();
                return;
            }

            _tenantCombo.Items.Clear();
            foreach (Tenant t in tenants)
            {
                _tenantCombo.Items.Add(t);
            }
            _tenantCombo.DisplayMember = "Name";
            _tenantCombo.Enabled = true;
            if (_tenantCombo.Items.Count > 0)
            {
                _tenantCombo.SelectedIndex = 0;
            }

            _confirmTenantButton.Tag = session;
        }

        private void OnConfirmTenantClicked()
        {
            var session = _confirmTenantButton.Tag as Session;
            var tenant = _tenantCombo.SelectedItem as Tenant;
            if (session == null || tenant == null)
            {
                _tenantErrorLabel.Text = "یک شرکت انتخاب کنید.";
                return;
            }

            Session updated = AuthService.SelectTenant(session, tenant.Id, tenant.Name);
            EnterWorkingState(updated);
        }

        private void OnSwitchTenantClicked()
        {
            Session session = AuthService.GetSession();
            if (session != null)
            {
                _ = ShowTenantSectionAsync(session);
            }
        }

        private void OnLogoutClicked()
        {
            AuthService.Logout();
            Teardown();
            _queue = null;
            ShowAuthSection();
        }

        private void EnterWorkingState(Session session)
        {
            HideAllSections();
            _workingPanel.Visible = true;
            _activeTenantLabel.Text = session.TenantName ?? "—";

            _queue = new ValidationQueue(
                _app,
                AuthService.GetSession,
                OnQueueStateChanged,
                OnErrorsReceived,
                OnAuthExpired);

            HookSheetChangeEvent();

            // Validate once immediately so data already in the sheet gets checked.
            _queue.TriggerNow();
        }

        private void HookSheetChangeEvent()
        {
            if (_sheetChangeHooked)
            {
                return; // already hooked
            }
            // += / -= with a plain method group lets the compiler infer the
            // exact Interop delegate type for us (it differs slightly across
            // Excel PIA versions), as long as the signature below matches.
            _app.SheetChange += App_SheetChange;
            _sheetChangeHooked = true;
        }

        private void App_SheetChange(object sh, Excel.Range target)
        {
            _queue?.MarkDirty();
        }

        private void Teardown()
        {
            if (_sheetChangeHooked)
            {
                _app.SheetChange -= App_SheetChange;
                _sheetChangeHooked = false;
            }
        }

        private void OnAuthExpired()
        {
            AuthService.Logout();
            Teardown();
            _queue = null;
            ShowAuthSection();
            _authErrorLabel.Text = "نشست شما منقضی شده است، دوباره وارد شوید.";
        }

        // ---------- status ----------

        private void OnQueueStateChanged(QueueState state)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => OnQueueStateChanged(state)));
                return;
            }

            switch (state.Status)
            {
                case ConnectivityStatus.Sending:
                    _statusLabel.Text = "در حال ارسال...";
                    _statusPanel.BackColor = WarnRowBg;
                    _statusLabel.ForeColor = WarnColor;
                    break;
                case ConnectivityStatus.Ok:
                    _statusLabel.Text = "به‌روز";
                    _statusPanel.BackColor = Color.FromArgb(0xE4, 0xEE, 0xEA);
                    _statusLabel.ForeColor = OkColor;
                    break;
                case ConnectivityStatus.Offline:
                    _statusLabel.Text = "عدم دسترسی به سرور - تلاش مجدد خودکار";
                    _statusPanel.BackColor = ErrorRowBg;
                    _statusLabel.ForeColor = ErrorColor;
                    break;
                case ConnectivityStatus.Error:
                    _statusLabel.Text = state.LastErrorMessage ?? "خطا";
                    _statusPanel.BackColor = ErrorRowBg;
                    _statusLabel.ForeColor = ErrorColor;
                    break;
                default:
                    _statusLabel.Text = "آماده";
                    _statusPanel.BackColor = Color.FromArgb(0xE4, 0xEE, 0xEA);
                    _statusLabel.ForeColor = OkColor;
                    break;
            }

            _pendingLabel.Text = state.PendingRowCount > 0 ? $"{state.PendingRowCount} سطر در حال بررسی" : string.Empty;
        }

        // ---------- results ----------

        private void OnErrorsReceived(string worksheetName, List<ValidationError> errors)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => OnErrorsReceived(worksheetName, errors)));
                return;
            }

            _lastWorksheetName = worksheetName;
            _resultsList.Items.Clear();

            if (errors == null || errors.Count == 0)
            {
                _resultsList.Visible = false;
                _emptyLabel.Visible = true;
                return;
            }

            _emptyLabel.Visible = false;
            _resultsList.Visible = true;

            foreach (var err in errors.OrderBy(e => e.SheetRowIndex))
            {
                var item = new ListViewItem(err.SheetRowIndex.ToString())
                {
                    Tag = err.SheetRowIndex,
                    BackColor = err.Severity == ValidationSeverity.Warning ? WarnRowBg : ErrorRowBg
                };
                item.SubItems.Add(err.Column ?? string.Empty);
                item.SubItems.Add(err.Message ?? string.Empty);
                _resultsList.Items.Add(item);
            }
        }

        private void OnResultActivated()
        {
            if (_resultsList.SelectedItems.Count == 0)
            {
                return;
            }
            int rowIndex = (int)_resultsList.SelectedItems[0].Tag;
            ExcelService.GoToRow(_app, _lastWorksheetName, rowIndex);
        }
    }
}
