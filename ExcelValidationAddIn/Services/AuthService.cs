using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using ExcelValidationAddIn.Models;

namespace ExcelValidationAddIn.Services
{
    public static class AuthService
    {
        public static Session GetSession()
        {
            return SettingsStore.Load();
        }

        /// <summary>
        /// Opens the accounting system's normal login page in the user's
        /// default browser (exactly the same login UI they already use day
        /// to day), waits for it to redirect back to a local loopback
        /// address with a one-time code, then exchanges that code for a JWT.
        /// Throws ApiException (timeout/network/auth) or
        /// OperationCanceledException if <paramref name="cancellationToken"/>
        /// is cancelled while waiting.
        /// </summary>
        public static async Task<Session> LoginAsync(CancellationToken cancellationToken)
        {
            int port = LoopbackListener.GetFreeTcpPort();
            string redirectUri = $"http://127.0.0.1:{port}/callback/";
            string state = Guid.NewGuid().ToString("N");

            string loginUrl =
                $"{ApiService.ApiBaseUrl}/desktop-login?returnUrl={Uri.EscapeDataString(redirectUri)}&state={state}";

            try
            {
                Process.Start(new ProcessStartInfo(loginUrl) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                throw new ApiException("امکان باز کردن مرورگر نبود: " + ex.Message, ApiErrorKind.Network);
            }

            LoopbackCallbackResult callback = await LoopbackListener
                .WaitForCallbackAsync(redirectUri, TimeSpan.FromMinutes(5), cancellationToken);

            if (!string.IsNullOrEmpty(callback.Error))
            {
                throw new ApiException("ورود در مرورگر ناموفق بود.", ApiErrorKind.Auth);
            }
            if (string.IsNullOrEmpty(callback.State) || callback.State != state)
            {
                throw new ApiException("پاسخ نامعتبر از سرور ورود دریافت شد.", ApiErrorKind.Auth);
            }
            if (string.IsNullOrEmpty(callback.Code))
            {
                throw new ApiException("کد ورود از سرور دریافت نشد.", ApiErrorKind.Auth);
            }

            LoginResponse response = await ApiService.ExchangeCodeAsync(callback.Code);

            var session = new Session
            {
                Token = response.Token,
                ExpiresAtUtc = response.ExpiresAt,
                TenantId = null,
                TenantName = null
            };
            SettingsStore.Save(session);
            return session;
        }

        public static void Logout()
        {
            SettingsStore.Clear();
        }

        public static Task<List<Tenant>> FetchTenantsAsync(string token)
        {
            return ApiService.FetchTenantsAsync(token);
        }

        public static Session SelectTenant(Session session, string tenantId, string tenantName)
        {
            session.TenantId = tenantId;
            session.TenantName = tenantName;
            SettingsStore.Save(session);
            return session;
        }
    }
}
