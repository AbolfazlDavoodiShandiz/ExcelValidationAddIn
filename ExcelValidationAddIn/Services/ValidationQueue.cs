using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelValidationAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelValidationAddIn.Services
{
    /// <summary>
    /// Watches for "dirty" signals from worksheet changes, waits for a pause
    /// in typing, then reads the whole sheet fresh and sends ONE batch
    /// request. Network failures never propagate back into Excel - they only
    /// update status and schedule a retry with exponential backoff.
    /// </summary>
    public class ValidationQueue
    {
        private const int DebounceMs = 3000;
        private const int InitialBackoffMs = 5000;
        private const int MaxBackoffMs = 60000;

        private readonly Excel.Application _app;
        private readonly Func<Session> _getSession;
        private readonly Action<QueueState> _onStateChange;
        private readonly Action<string, List<ValidationError>> _onErrors;
        private readonly Action _onAuthExpired;

        private readonly Timer _debounceTimer;
        private readonly Timer _retryTimer;

        private int _backoffMs = InitialBackoffMs;
        private bool _dirty;
        private bool _inFlight;
        private QueueState _state = new QueueState();

        public ValidationQueue(
            Excel.Application app,
            Func<Session> getSession,
            Action<QueueState> onStateChange,
            Action<string, List<ValidationError>> onErrors,
            Action onAuthExpired)
        {
            _app = app;
            _getSession = getSession;
            _onStateChange = onStateChange;
            _onErrors = onErrors;
            _onAuthExpired = onAuthExpired;

            // Windows Forms Timer, not System.Threading.Timer: it fires back
            // on this (STA/UI) thread, which is required since flushing
            // touches Excel COM objects.
            _debounceTimer = new Timer { Interval = DebounceMs };
            _debounceTimer.Tick += (s, e) =>
            {
                _debounceTimer.Stop();
                _ = FlushAsync();
            };

            _retryTimer = new Timer();
            _retryTimer.Tick += (s, e) =>
            {
                _retryTimer.Stop();
                _ = FlushAsync();
            };
        }

        /// <summary>Call from the Application.SheetChange handler. Cheap and non-blocking.</summary>
        public void MarkDirty()
        {
            _dirty = true;
            SetState(Clone(_state, status: _state.Status == ConnectivityStatus.Offline
                ? ConnectivityStatus.Offline
                : ConnectivityStatus.Idle));

            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        /// <summary>Call from the "Validate Now" button. Skips the debounce wait.</summary>
        public void TriggerNow()
        {
            _debounceTimer.Stop();
            _dirty = true;
            _ = FlushAsync();
        }

        private void SetState(QueueState state)
        {
            _state = state;
            _onStateChange?.Invoke(_state);
        }

        private async Task FlushAsync()
        {
            if (_inFlight || !_dirty)
            {
                return;
            }

            Session session = _getSession();
            if (session == null || string.IsNullOrEmpty(session.TenantId))
            {
                // Nothing to do without a session - the task pane already
                // prompts the user to sign in / pick a tenant in that case.
                return;
            }

            _inFlight = true;
            _dirty = false;
            SetState(Clone(_state, status: ConnectivityStatus.Sending));

            try
            {
                SheetSnapshot snapshot = ExcelService.ReadActiveSheet(_app);
                if (snapshot.Rows.Count == 0)
                {
                    SetState(Clone(_state, status: ConnectivityStatus.Idle, pendingRowCount: 0));
                    return;
                }

                SetState(Clone(_state, status: ConnectivityStatus.Sending, pendingRowCount: snapshot.Rows.Count));

                var request = new BatchValidationRequest { Rows = snapshot.Rows };

                // ConfigureAwait(true): resume on this STA thread, since the
                // rendering step right after touches Excel COM objects again.
                BatchValidationResponse result = await ApiService.ValidateBatchAsync(session, request).ConfigureAwait(true);

                ExcelService.RenderValidationResults(_app, snapshot.WorksheetName, snapshot.Headers.Count, result.Errors);
                _onErrors?.Invoke(snapshot.WorksheetName, result.Errors);

                _backoffMs = InitialBackoffMs;
                SetState(new QueueState
                {
                    Status = ConnectivityStatus.Ok,
                    PendingRowCount = 0,
                    LastSuccessAtUtc = DateTime.UtcNow,
                    LastErrorMessage = null
                });
            }
            catch (ApiException apiEx)
            {
                if (apiEx.Kind == ApiErrorKind.Auth)
                {
                    SetState(Clone(_state, status: ConnectivityStatus.Error, lastErrorMessage: "نشست منقضی شده - دوباره وارد شوید."));
                    _onAuthExpired?.Invoke();
                }
                else
                {
                    string message = apiEx.Kind == ApiErrorKind.Timeout
                        ? "سرور اعتبارسنجی پاسخ نداد. تلاش مجدد خودکار انجام می‌شود."
                        : apiEx.Kind == ApiErrorKind.Network
                            ? "اتصال به سرور برقرار نشد. تلاش مجدد خودکار انجام می‌شود."
                            : "اعتبارسنجی ناموفق بود. تلاش مجدد خودکار انجام می‌شود.";

                    SetState(Clone(_state, status: ConnectivityStatus.Offline, lastErrorMessage: message));

                    // Keep the data marked dirty and retry with backoff - a
                    // long outage must not spam the server or the user.
                    _dirty = true;
                    ScheduleRetry();
                }
            }
            catch (Exception)
            {
                // Defensive catch-all: a bug here must never propagate into
                // Excel's edit path or crash the host application.
                SetState(Clone(_state, status: ConnectivityStatus.Offline, lastErrorMessage: "خطای غیرمنتظره در اعتبارسنجی."));
                _dirty = true;
                ScheduleRetry();
            }
            finally
            {
                _inFlight = false;
            }
        }

        private void ScheduleRetry()
        {
            _retryTimer.Stop();
            _retryTimer.Interval = _backoffMs;
            _backoffMs = Math.Min(_backoffMs * 2, MaxBackoffMs);
            _retryTimer.Start();
        }

        private static QueueState Clone(
            QueueState source,
            ConnectivityStatus? status = null,
            int? pendingRowCount = null,
            string lastErrorMessage = null)
        {
            return new QueueState
            {
                Status = status ?? source.Status,
                PendingRowCount = pendingRowCount ?? source.PendingRowCount,
                LastSuccessAtUtc = source.LastSuccessAtUtc,
                LastErrorMessage = lastErrorMessage ?? source.LastErrorMessage
            };
        }
    }
}
