using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelValidationAddIn.Models;
using Newtonsoft.Json;

namespace ExcelValidationAddIn.Services
{
    public enum ApiErrorKind
    {
        Network,
        Timeout,
        Auth,
        Server
    }

    public class ApiException : Exception
    {
        public ApiErrorKind Kind { get; }

        public ApiException(string message, ApiErrorKind kind) : base(message)
        {
            Kind = kind;
        }
    }

    public static class ApiService
    {
        // TODO: point this at the real accounting-system API before shipping.
        public const string ApiBaseUrl = "http://localhost:5080";

        private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(15);

        private static readonly HttpClient Http = new HttpClient
        {
            BaseAddress = new Uri(ApiBaseUrl)
        };

        public static async Task<LoginResponse> ExchangeCodeAsync(string code)
        {
            var body = new { code };
            var content = new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json");

            using (var cts = new CancellationTokenSource(RequestTimeout))
            {
                HttpResponseMessage response;
                try
                {
                    response = await Http.PostAsync("/api/auth/exchange", content, cts.Token).ConfigureAwait(false);
                }
                catch (TaskCanceledException)
                {
                    throw new ApiException("درخواست تبادل کد ورود با تایم‌اوت مواجه شد.", ApiErrorKind.Timeout);
                }
                catch (HttpRequestException)
                {
                    throw new ApiException("اتصال به سرور حسابداری برقرار نشد.", ApiErrorKind.Network);
                }

                if (!response.IsSuccessStatusCode)
                {
                    throw new ApiException("کد ورود نامعتبر یا منقضی شده است.", ApiErrorKind.Auth);
                }

                string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                return JsonConvert.DeserializeObject<LoginResponse>(json);
            }
        }

        public static async Task<System.Collections.Generic.List<Tenant>> FetchTenantsAsync(string token)
        {
            using (var request = new HttpRequestMessage(HttpMethod.Get, "/api/tenants"))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                using (var cts = new CancellationTokenSource(RequestTimeout))
                {
                    HttpResponseMessage response;
                    try
                    {
                        response = await Http.SendAsync(request, cts.Token).ConfigureAwait(false);
                    }
                    catch (TaskCanceledException)
                    {
                        throw new ApiException("دریافت لیست شرکت‌ها با تایم‌اوت مواجه شد.", ApiErrorKind.Timeout);
                    }
                    catch (HttpRequestException)
                    {
                        throw new ApiException("اتصال به سرور حسابداری برقرار نشد.", ApiErrorKind.Network);
                    }

                    if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                        response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                    {
                        throw new ApiException("نشست شما معتبر نیست.", ApiErrorKind.Auth);
                    }

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new ApiException($"سرور خطای HTTP {(int)response.StatusCode} بازگرداند.", ApiErrorKind.Server);
                    }

                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return JsonConvert.DeserializeObject<System.Collections.Generic.List<Tenant>>(json);
                }
            }
        }

        public static async Task<BatchValidationResponse> ValidateBatchAsync(Session session, BatchValidationRequest batch)
        {
            if (string.IsNullOrEmpty(session.TenantId))
            {
                throw new ApiException("هیچ شرکتی انتخاب نشده است.", ApiErrorKind.Auth);
            }

            using (var request = new HttpRequestMessage(HttpMethod.Post, "/api/accounting-documents/validate-batch"))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", session.Token);
                request.Headers.Add("X-Tenant-Id", session.TenantId);
                request.Content = new StringContent(JsonConvert.SerializeObject(batch), Encoding.UTF8, "application/json");

                using (var cts = new CancellationTokenSource(RequestTimeout))
                {
                    HttpResponseMessage response;
                    try
                    {
                        response = await Http.SendAsync(request, cts.Token).ConfigureAwait(false);
                    }
                    catch (TaskCanceledException)
                    {
                        throw new ApiException("درخواست اعتبارسنجی با تایم‌اوت مواجه شد.", ApiErrorKind.Timeout);
                    }
                    catch (HttpRequestException)
                    {
                        // Covers "server is down" / DNS / connection-refused cases.
                        throw new ApiException("امکان اتصال به سرور اعتبارسنجی نبود.", ApiErrorKind.Network);
                    }

                    if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                        response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                    {
                        throw new ApiException("نشست شما منقضی شده یا برای این شرکت مجاز نیست.", ApiErrorKind.Auth);
                    }

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new ApiException($"سرور خطای HTTP {(int)response.StatusCode} بازگرداند.", ApiErrorKind.Server);
                    }

                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return JsonConvert.DeserializeObject<BatchValidationResponse>(json);
                }
            }
        }
    }
}
