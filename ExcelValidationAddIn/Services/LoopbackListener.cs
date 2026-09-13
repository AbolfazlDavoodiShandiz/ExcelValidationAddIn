using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelValidationAddIn.Services
{
    public class LoopbackCallbackResult
    {
        public string Code { get; set; }
        public string State { get; set; }
        public string Error { get; set; }
    }

    /// <summary>
    /// Starts a short-lived HTTP listener on 127.0.0.1 and waits for the
    /// single browser redirect that completes a "login in the system
    /// browser" flow (the same pattern GitHub Desktop / VS Code / MSAL use).
    /// Binding to a loopback address specifically does NOT require
    /// administrator rights or a netsh URL-ACL reservation on Windows.
    /// </summary>
    public static class LoopbackListener
    {
        public static int GetFreeTcpPort()
        {
            var tcp = new TcpListener(IPAddress.Loopback, 0);
            tcp.Start();
            int port = ((IPEndPoint)tcp.LocalEndpoint).Port;
            tcp.Stop();
            return port;
        }

        /// <summary>
        /// Listens on <paramref name="redirectUri"/> (must end with '/') until
        /// one request arrives, <paramref name="timeout"/> elapses, or
        /// <paramref name="cancellationToken"/> is cancelled (e.g. the user
        /// clicked "Cancel" in the task pane while waiting).
        /// </summary>
        public static async Task<LoopbackCallbackResult> WaitForCallbackAsync(
            string redirectUri,
            TimeSpan timeout,
            CancellationToken cancellationToken)
        {
            using (var listener = new HttpListener())
            {
                listener.Prefixes.Add(redirectUri);

                try
                {
                    listener.Start();
                }
                catch (HttpListenerException ex)
                {
                    throw new ApiException(
                        "امکان راه‌اندازی سرور محلی برای دریافت نتیجه ورود نبود: " + ex.Message,
                        ApiErrorKind.Network);
                }

                using (cancellationToken.Register(() => SafeStop(listener)))
                {
                    Task<HttpListenerContext> getContextTask = listener.GetContextAsync();
                    Task timeoutTask = Task.Delay(timeout, CancellationToken.None);
                    Task completed = await Task.WhenAny(getContextTask, timeoutTask).ConfigureAwait(false);

                    if (completed == timeoutTask)
                    {
                        SafeStop(listener);
                        throw new ApiException("ورود از طریق مرورگر با تایم‌اوت مواجه شد.", ApiErrorKind.Timeout);
                    }

                    if (cancellationToken.IsCancellationRequested)
                    {
                        throw new OperationCanceledException(cancellationToken);
                    }

                    HttpListenerContext context;
                    try
                    {
                        context = await getContextTask.ConfigureAwait(false);
                    }
                    catch (Exception) when (cancellationToken.IsCancellationRequested)
                    {
                        throw new OperationCanceledException(cancellationToken);
                    }

                    var result = new LoopbackCallbackResult
                    {
                        Code = context.Request.QueryString["code"],
                        State = context.Request.QueryString["state"],
                        Error = context.Request.QueryString["error"]
                    };

                    WriteResponsePage(context, success: result.Error == null);
                    SafeStop(listener);
                    return result;
                }
            }
        }

        private static void WriteResponsePage(HttpListenerContext context, bool success)
        {
            string title = success ? "ورود موفقیت‌آمیز بود" : "ورود ناموفق بود";
            string body = success
                ? "می‌توانید این تب را ببندید و به Excel بازگردید."
                : "لطفاً این تب را ببندید و در افزونه اکسل دوباره تلاش کنید.";

            string html =
                "<html><head><meta charset='utf-8'></head>" +
                "<body style='font-family:Segoe UI,Tahoma,sans-serif;text-align:center;padding:60px;color:#1c2530'>" +
                $"<h2>{title}</h2><p>{body}</p></body></html>";

            byte[] buffer = Encoding.UTF8.GetBytes(html);
            try
            {
                context.Response.ContentType = "text/html; charset=utf-8";
                context.Response.ContentLength64 = buffer.Length;
                context.Response.OutputStream.Write(buffer, 0, buffer.Length);
                context.Response.OutputStream.Close();
            }
            catch
            {
                // The browser tab may already be closed - nothing to do.
            }
        }

        private static void SafeStop(HttpListener listener)
        {
            try
            {
                if (listener.IsListening)
                {
                    listener.Stop();
                }
            }
            catch
            {
                // Best-effort cleanup only.
            }
        }
    }
}
