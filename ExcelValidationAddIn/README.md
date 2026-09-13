# Accounting Document Validation - Excel VSTO Add-in (C#)

A native Excel Add-in (VSTO, C#, .NET Framework 4.8) with the exact same
behavior as the web/Office.js version: validates accounting-document rows
against the accounting system's API while the user fills the raw Excel
template, highlights bad rows, adds a cell comment per row, and lists errors
in a task pane.

## Producing the real Setup.exe (ClickOnce)

Once the project builds in Visual Studio:

1. Right-click the project ▸ **Properties ▸ Publish**.
2. Set **Publish Location** to a folder (e.g. `C:\Builds\ExcelValidationAddIn\`)
   or a network/file share your users can reach.
3. Click **Finish**, then **Publish Now**.
4. Visual Studio produces a folder containing `setup.exe`, the `.vsto`
   application manifest, and the add-in files. Users run `setup.exe`, which
   installs the VSTO runtime if needed, installs the add-in, and registers
   it with Excel - no manual DLL registration required.
5. For real deployment, sign the ClickOnce manifest with a code-signing
   certificate (**Properties ▸ Signing**) - unsigned add-ins trigger extra
   Windows/Office security prompts on your users' machines.

Alternative (no Visual Studio IDE, just Build Tools): install "Visual
Studio Build Tools" with the Office/SharePoint workload, then from a
**Developer Command Prompt**:
```
msbuild ExcelValidationAddIn.sln /t:Publish /p:Configuration=Release
```

## What's implemented, and where

| Requirement | Where |
|---|---|
| Browser-based login (like GitHub Desktop) | `Services/AuthService.LoginAsync` opens the accounting system's normal login page in the **system's default browser** and starts a short-lived local listener on `127.0.0.1` (`Services/LoopbackListener.cs`) to catch the redirect after login - see "Browser-based login flow" below for the full sequence and the small backend piece it needs. |
| Multi-tenant, tenant sent per request | `UI/ValidationPaneControl.cs` fetches `/api/tenants` after login and lets the user pick one; `Services/ApiService.ValidateBatchAsync` sends it as the `X-Tenant-Id` header. |
| No request per keystroke, batched | `Services/ValidationQueue.cs` debounces `Application.SheetChange` (3s) and always re-reads the current sheet on flush. |
| Server down/timeout must not block data entry | `ApiService` uses a per-request timeout (`CancellationTokenSource`); `ValidationQueue` catches every failure, never throws back into the sheet-change handler, and retries with exponential backoff. Excel's own editing is never awaited on the network call. |
| Error rows highlighted + comment | `Services/ExcelService.RenderValidationResults` fills flagged rows and adds a native cell comment per row; clears only the marks it made on the previous run. |
| Errors shown in a task pane | `UI/ValidationPaneControl.cs`, hosted via `CustomTaskPanes.Add` in `ThisAddIn.cs`; double-click a result to jump to that cell. A ribbon button (`Ribbon/Ribbon.xml` + `Ribbon/ValidationRibbon.cs`) toggles the pane. |

## Browser-based login flow

Clicking "ورود به سیستم" does **not** open any form inside the add-in.
Instead, it follows the same loopback pattern GitHub Desktop, VS Code,
`az login`, Google Cloud SDK, etc. use:

1. `AuthService.LoginAsync` picks a free local port and starts an
   `HttpListener` bound to `http://127.0.0.1:{port}/callback/`
   (`Services/LoopbackListener.cs`). Binding to a loopback address
   specifically does **not** require administrator rights or a `netsh`
   URL-ACL reservation on Windows.
2. It opens the user's **default system browser** at:
   `{ApiBaseUrl}/desktop-login?returnUrl=http://127.0.0.1:{port}/callback/&state={random}`
3. The user logs in on the accounting system's normal, existing web login
   page - nothing new for them to learn.
4. On success, the server redirects the browser to
   `returnUrl?code={oneTimeCode}&state={random}`. The add-in's local
   listener catches this one request, shows a small "you can close this
   tab" page, and stops listening.
5. The add-in then calls `POST /api/auth/exchange` **directly** (not
   through the browser) with `{ code }` to get the real JWT. This keeps
   the JWT out of the browser's address bar/history - only a short-lived,
   single-use code ever touches the browser.
6. The task pane shows a "در حال انتظار برای ورود در مرورگر..." state with
   a Cancel button while waiting (5-minute timeout either way).

This needs **one small addition on the ASP.NET 6 side**: a
`/desktop-login` endpoint that piggybacks on your *existing* login
page/logic - it doesn't need its own login form, it just needs to know
where to send the browser back to once your existing login succeeds.
Rough shape:

```csharp
[HttpGet("/desktop-login")]
public IActionResult DesktopLogin([FromQuery] string returnUrl, [FromQuery] string state)
{
    // If there's no authenticated browser session yet, show your EXISTING
    // login page/flow, remembering returnUrl+state (TempData, a signed
    // cookie, or chaining them through your login POST) until the user is
    // authenticated.
    if (!User.Identity.IsAuthenticated)
    {
        return RedirectToYourExistingLoginPage(returnUrl, state);
    }

    // Already authenticated: mint a short-lived, single-use code instead
    // of handing out the JWT directly in the URL.
    string code = _desktopLoginCodes.IssueOneTimeCode(User);
    return Redirect($"{returnUrl}?code={code}&state={state}");
}

[HttpPost("/api/auth/exchange")]
public IActionResult ExchangeCode([FromBody] ExchangeCodeRequest request)
{
    ClaimsPrincipal principal = _desktopLoginCodes.RedeemOneTimeCode(request.Code); // single use, short TTL (~2 min)
    if (principal == null)
    {
        return Unauthorized();
    }
    var (token, expiresAt) = _jwtIssuer.IssueToken(principal); // reuse your existing JWT issuance
    return Ok(new { token, expiresAt });
}
```

`_desktopLoginCodes` can be as simple as an `IMemoryCache` mapping a
random code to the authenticated user's identity, short TTL, redeemed once.

## Configuration before building

- `Services/ApiService.cs` - set `ApiBaseUrl` to the real accounting API
  (also used to build the `/desktop-login` URL).
- Make sure `/desktop-login` and `/api/auth/exchange` exist on the server
  (see above) before testing login end to end.
- App icon / manifest branding: Project Properties ▸ Application.

## Backend contract expected

**`GET /desktop-login?returnUrl=...&state=...`** - see "Browser-based login
flow" above; wraps your existing login page, ends with a redirect to
`returnUrl?code=...&state=...`.

**`POST /api/auth/exchange`**
```json
// request
{ "code": "the one-time code from the redirect" }
// response 200
{ "token": "eyJ...", "expiresAt": "2026-09-08T20:00:00Z" }
```

**`GET /api/tenants`** (Authorization: Bearer &lt;token&gt;)
```json
[{ "id": "t-001", "name": "Acme Co." }, { "id": "t-002", "name": "Beta Ltd." }]
```

**`POST /api/accounting-documents/validate-batch`**
(Authorization: Bearer &lt;token&gt;, X-Tenant-Id: &lt;tenantId&gt;)
```json
{
  "rows": [
    { "sheetRowIndex": 2, "values": { "AccountCode": "1101", "Amount": 500000 } },
    { "sheetRowIndex": 3, "values": { "AccountCode": "", "Amount": -10 } }
  ]
}
// response - only rows with problems need to appear
{
  "errors": [
    { "sheetRowIndex": 3, "column": "AccountCode", "message": "کد حساب الزامی است.", "severity": "Error" }
  ]
}
```

Ideally the validate-batch endpoint reuses the exact same row-mapping/
validation service the upload flow already runs, just skipping the
file-parsing step.

## Known limitations / before production

- **COM lifetime**: `ExcelService.cs` releases every COM object it touches
  (`Marshal.ReleaseComObject`). If you extend it, keep that discipline -
  it's the most common source of "Excel won't close" bugs in VSTO add-ins.
- **Large sheets**: `RenderValidationResults` does simple per-row COM calls;
  fine for typical accounting-document batches, but for very large sheets
  consider batching the writes further.
- **Office bitness**: the Interop references target Office 15.0 (2013)
  type libraries, which are forward-compatible with 2016/2019/365 - just
  make sure the target machine's installed Office bitness (32/64-bit)
  matches how you publish the add-in (`Platform` in the Publish settings).
- **Token refresh**: on 401/403 the pane forces a re-login. Add a refresh
  token flow in `AuthService`/`ApiService` if the API supports one.
- **Loopback listener**: `LoopbackListener` binds only to `127.0.0.1`, so
  it never accepts connections from the network and does not trigger a
  Windows Firewall prompt (those only fire for inbound connections from
  other machines). It also stops itself immediately after the single
  callback request, on timeout (5 min), or on Cancel.
