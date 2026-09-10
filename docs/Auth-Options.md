# Authentication options without a service principal

ImpactIQ v3 signs in as a **user**. Service-principal support lives in the separate
[Service Principal Edition](https://github.com/BeSmarterWithData/ImpactIQ-ServicePrincipal); nothing here needs an app
registration - all modes use Microsoft's own pre-consented public clients (the same ones Azure PowerShell and the
Power BI PowerShell module use).

Contents: 1 modes at a glance - 2 `Auto` resolution - 3 GCC / sovereign endpoints - 4 each mode in detail -
5 token cache - 6 long runs - 7 MFA, Conditional Access and other things that break user auth - 8 security notes -
9 AADSTS quick reference.

## 1. Modes at a glance

| `-AuthMode` | How it signs in | Needs a human? | Survives MFA/CA? | Persists between runs? | Fabric token? | Typical use |
|---|---|---|---|---|---|---|
| `Interactive` | `Connect-PowerBIServiceAccount` (browser / WAM) via the MicrosoftPowerBIMgmt module, Fabric via `Connect-AzAccount` | every run | yes | no | best effort | the v2 experience; refused when headless |
| `DeviceCode` | pure HTTP OAuth 2.0 device-code flow; prints a code, a human enters it at `https://microsoft.com/devicelogin`; the **refresh token is cached (encrypted)** and reused silently | first run only (and after cache loss) | yes | yes (`State\auth\token-cache.json`) | yes, if the tenant offers Fabric (not in GCC) | **scheduled runs** - hosted agent with `IMPACTIQ_TOKEN_CACHE_KEY`, or self-hosted with DPAPI |
| `Credential` | OAuth 2.0 password grant (ROPC) with `-Credential` or `IMPACTIQ_USERNAME` / `IMPACTIQ_PASSWORD`; falls back to `Connect-PowerBIServiceAccount -Credential` on Windows when the module is installed | never | **no** | not needed | via the refresh token from the ROPC response | MFA-exempt service accounts |
| `AzContext` | reuses an existing Az PowerShell context (`Get-AzContext`), tokens via `Get-AzAccessToken -ResourceUrl` | once, to create the context (`Connect-AzAccount -UseDeviceAuthentication`) | yes | yes (Az's own MSAL cache under `%LOCALAPPDATA%\.IdentityService`, DPAPI) | yes, where Fabric exists | self-hosted agent / Task Scheduler on a box where Az is already set up; never auto-selected |
| `AccessToken` | `IMPACTIQ_PBI_TOKEN` (+ optional `IMPACTIQ_FABRIC_TOKEN`) minted by something else | n/a | n/a | no refresh at all; the run fails when the token expires (~1 h) | only if you supply one | one-off runs and tests |

## 2. `Auto` resolution (default)

`-AuthMode Auto` picks the first that applies:

1. `Credential` when `-Credential` is given **or** both `IMPACTIQ_USERNAME` and `IMPACTIQ_PASSWORD` are set.
2. `AccessToken` when `IMPACTIQ_PBI_TOKEN` is set.
3. `DeviceCode` when a token cache exists at the cache path (silent refresh) - even in an interactive console.
4. `AzContext` **only** when explicitly requested (`-AuthMode AzContext`); never auto-selected.
5. `Interactive` when the session is interactive (`Test-IQInteractive`: not `-NonInteractive`, `[Environment]::UserInteractive`,
   no `TF_BUILD`, no `CI`, console host present).
6. Otherwise `DeviceCode` (prints the code; a human completes it; the run then continues unattended for hours).

The resolved mode and the signed-in account appear in the log and in `manifest.json` (`"auth": "DeviceCode (cached) as user@agency.gov"`).

## 3. GCC and other sovereign clouds - the values that matter

The GCC (moderate) tenant lives in **commercial Entra ID**: the OAuth authority is `login.microsoftonline.com`, only the
Power BI resource/API hosts change. GCC High and DoD use `login.microsoftonline.us` and Azure Government. The v2 script
had the GCC High / DoD resource hosts in the wrong order (`analysis.high.usgovcloudapi.net`); v3 uses the values below
(Microsoft Learn "Embed content for national/regional clouds" and "Power BI for US Government customers").

| `-Environment` | REST `ApiPrefix` | OAuth authority | Token resource (`scope = <resource>/.default`) | XMLA prefix | Fabric API prefix | Az environment |
|---|---|---|---|---|---|---|
| `Public` | `https://api.powerbi.com` | `https://login.microsoftonline.com` | `https://analysis.windows.net/powerbi/api` | `powerbi://api.powerbi.com` | `https://api.fabric.microsoft.com` | `AzureCloud` |
| `USGov` (GCC) | `https://api.powerbigov.us` | **`https://login.microsoftonline.com`** | `https://analysis.usgovcloudapi.net/powerbi/api` | `powerbi://api.powerbigov.us` | `https://api.fabric.microsoft.us` (unverified - Fabric is not offered in GCC; calls degrade to empty) | `AzureCloud` |
| `USGovHigh` (GCC High) | `https://api.high.powerbigov.us` | `https://login.microsoftonline.us` | `https://high.analysis.usgovcloudapi.net/powerbi/api` | `powerbi://api.high.powerbigov.us` | `https://api.fabric.high.microsoft.us` (unverified; Microsoft's GCC High preview documents `https://highapi.fabric.microsoft.us` - use `-FabricApiPrefixOverride`) | `AzureUSGovernment` |
| `USGovMil` (DoD) | `https://api.mil.powerbigov.us` | `https://login.microsoftonline.us` | `https://mil.analysis.usgovcloudapi.net/powerbi/api` | `powerbi://api.mil.powerbigov.us` | `https://api.fabric.mil.microsoft.us` (unverified) | `AzureUSGovernment` |
| `China` | `https://api.powerbi.cn` | `https://login.chinacloudapi.cn` | `https://analysis.chinacloudapi.cn/powerbi/api` | `powerbi://api.powerbi.cn` | `https://api.fabric.microsoft.cn` | `AzureChinaCloud` |
| `Germany` | `https://api.powerbi.de` | `https://login.microsoftonline.de` | `https://analysis.cloudapi.de/powerbi/api` | `powerbi://api.powerbi.de` | `https://api.fabric.microsoft.de` | `AzureGermanCloud` |

Overrides: `-AuthorityOverride https://login.microsoftonline.com/<tenant-guid>` and `-FabricApiPrefixOverride <url>`.
Aliases `GCC`, `GCCHigh`, `DoD` map to `USGov`, `USGovHigh`, `USGovMil`.

GCC consequences:

* `Connect-AzAccount` for a GCC tenant uses the **default** `AzureCloud` environment (plus `-Tenant <guid>`), not
  `-Environment AzureUSGovernment` - that one is for GCC High / DoD.
* The Power BI PowerShell module (Interactive / Credential fallback) first calls the commercial discovery endpoint
  `https://api.powerbi.com/powerbi/globalservice/...` to learn the GCC hosts. A hardened agent that only allows
  `*.powerbigov.us` fails with `Failed to populate environments in settings`. DeviceCode / Credential (HTTP) modes only
  need `login.microsoftonline.com` and `api.powerbigov.us`.
* The device-login URL printed is `https://microsoft.com/devicelogin` for the commercial authority; for
  `login.microsoftonline.us` the `verification_uri` returned by the endpoint is printed (always trust the printed one).
* Fabric REST is not available in GCC moderate: `Get-IQToken -Resource Fabric` returns `$null`, Fabric collectors
  produce empty sheets, `getDefinition` fallbacks are skipped. No error, one Debug line per run.

## 4. Each mode in detail

### 4.1 DeviceCode (recommended for schedules)

What the tool does (`Initialize-IQAuth -Mode DeviceCode`):

1. `POST {Authority}/{TenantId}/oauth2/v2.0/devicecode` with `client_id` and
   `scope = "<PowerBIResource>/.default offline_access openid profile"`.
2. Logs the endpoint's `message` at **Warn** level: `DEVICE CODE SIGN-IN REQUIRED: To sign in, use a web browser to open
   the page https://microsoft.com/devicelogin and enter the code ABCD1234 to authenticate.` If
   `-DeviceCodeWebhookUrl` / `IMPACTIQ_DEVICECODE_WEBHOOK` is set, the same text is POSTed as `{ "text": "..." }` to
   the Teams/Slack incoming webhook (never throws; a webhook failure is logged and the run continues).
3. Polls `POST .../oauth2/v2.0/token` (`grant_type=urn:ietf:params:oauth:grant-type:device_code`) every `interval`
   seconds, handling `authorization_pending`, `slow_down` (+5 s), `expired_token` (after 15 min -> throw),
   `authorization_declined` / `bad_verification_code` (throw), and Conditional Access codes 50076/50079/53003/530036
   ("blocked by Conditional Access" throw).
4. Stores `access_token` / `refresh_token` / expiry per resource; redeems the same refresh token for the Fabric resource
   (`<FabricApiPrefix>/.default`) - if that fails, Fabric tokens are `$null` for the run.
5. Persists the newest refresh token with `Save-IQTokenCache` after **every** successful refresh (refresh tokens rotate).

One-time setup (nothing to install):

```powershell
# self-hosted agent / Task Scheduler: run once AS the account the schedule will use
.\ImpactIQ.ps1 -BaseFolder 'C:\ImpactIQ' -NonInteractive -Environment USGov -AuthMode DeviceCode -Stages Inventory -WorkspaceName 'Finance'
# hosted agent: set IMPACTIQ_TOKEN_CACHE_KEY as a secret variable first, then run the pipeline once by hand
```

`-TenantId` defaults to `organizations`; set it to your tenant GUID (`IMPACTIQ_TENANT_ID`) when the account is a guest
in other tenants or the sign-in lands in the wrong tenant. `-ClientId` defaults to Azure PowerShell's first-party
public client `1950a258-227b-4e31-a9cf-717495945fc2` (pre-consented for Power BI in every cloud); Azure CLI's
`04b07795-8ddb-461a-bbee-02f9e1bf7b46` is the documented alternative. **A cached refresh token is bound to the client id
that issued it** - changing `-ClientId` invalidates the cache (the tool detects this and starts a fresh device-code flow).

When the cached refresh token is rejected mid-run (`invalid_grant`, `interaction_required`), the tool logs a Warn,
posts to the webhook and starts a **new** device-code flow instead of failing - a human can rescue an unattended run.
If nobody does within 15 minutes the run ends with exit 1 and is resumed by the next scheduled run.

### 4.2 Credential (ROPC)

`POST .../oauth2/v2.0/token` with `grant_type=password&username=...&password=...&client_id=...&scope=...`. The password
comes from `-Credential` or `IMPACTIQ_PASSWORD`, is converted to a `SecureString` immediately and only decrypted inside
the form body. It is never logged and never written to the manifest.

Requirements for the account: cloud-only (password hash sync or pass-through auth - **not** AD FS federated unless
`AllowCloudPasswordValidation` is set), **not** subject to MFA or a Conditional Access policy that requires MFA /
compliant device / auth strength, not passwordless-only, not a guest, password without leading/trailing spaces, a Pro
or PPU license, and (for GCC) `-TenantId` set when `organizations` does not resolve.

Failure mapping (`AADSTS` codes in the token response body):

| Codes | Message from the tool | What to do |
|---|---|---|
| `50076`, `50079`, `53003`, `65001`, `50074`, `530036` | "MFA/Conditional Access blocks password auth for this account; use DeviceCode with a cached refresh token or exempt this account" (fatal, no fallback) | switch to DeviceCode, or exempt the account |
| `50126` | invalid username or password | check the secret variable; watch for trailing spaces |
| `50034`, `90002`, `50059` | user / tenant not found | wrong `IMPACTIQ_USERNAME` or `IMPACTIQ_TENANT_ID`, or wrong `-Environment` (GCC High/DoD use `login.microsoftonline.us`) |
| `50053` | account locked (smart lockout) | wait, then fix the password |
| `50055`, `50056`, `50057` | password expired / no password / account disabled | identity team |
| `50155` | device authentication required | the account is subject to device-based CA - use DeviceCode from a compliant device or exempt |
| anything else | falls back to `Connect-PowerBIServiceAccount -Environment <Env> -Credential` when running on Windows with `MicrosoftPowerBIMgmt.Profile` installed (no Fabric token in that path); otherwise fatal | install the module on the agent (`Install-Module MicrosoftPowerBIMgmt.Profile -Scope CurrentUser`) or switch modes |

Microsoft deprecates ROPC (MSAL `AcquireTokenByUsernamePassword` deprecated since 4.74; mandatory-MFA program).
Treat mode A as a stop-gap.

### 4.3 AzContext

`Import-Module Az.Accounts`; requires an existing context (`Get-AzContext`), else throws with the setup instructions:

```powershell
# once, as the account the schedule uses (GCC = commercial Entra -> default AzureCloud environment):
Connect-AzAccount -UseDeviceAuthentication -Tenant <tenant-guid>            # GCC / Public
Connect-AzAccount -UseDeviceAuthentication -Environment AzureUSGovernment   # GCC High / DoD
Enable-AzContextAutosave -Scope CurrentUser
# verify in a NEW PowerShell session (this is what the scheduled run sees):
Get-AzContext; Get-AzAccessToken -ResourceUrl 'https://analysis.usgovcloudapi.net/powerbi/api'
```

Tokens: `Get-AzAccessToken -ResourceUrl <PowerBIResource>` (and `<FabricApiPrefix>` for Fabric); the `SecureString`
token of Az 14+ is unwrapped exactly as the v2 script did. Az never signs in on its own inside ImpactIQ - when the
context is gone (90 idle days, password change, `Clear-AzContext`), the run fails fast with the instructions above.
Az's MSAL cache is DPAPI-protected per Windows user: bootstrap **as the agent's service account** (a manual pipeline run
on that agent is the practical way). `Update-AzConfig -LoginExperienceV2 Off` avoids the interactive subscription
picker for accounts without Azure subscriptions.

### 4.4 AccessToken

`IMPACTIQ_PBI_TOKEN` (bearer token for the Power BI resource, `Bearer ` prefix optional), optional
`IMPACTIQ_FABRIC_TOKEN`. The tool logs the expiry (`exp` claim) at Warn on start, throws immediately if it is already
expired, and throws with a clear message when it expires mid-run. Useful to test a pipeline with a token from
`Get-PowerBIAccessToken -AsString` or `az account get-access-token --resource https://analysis.usgovcloudapi.net/powerbi/api`.

### 4.5 Interactive

The v2 flow moved verbatim: `Connect-PowerBIServiceAccount -Environment <Env>` (omitted for Public) with the same
two-attempt retry, `Get-PowerBIAccessToken` for tokens, Fabric best-effort through `Connect-AzAccount -Environment
<AzEnvironment> -Scope Process -SkipContextPopulation` + `Get-AzAccessToken`. Refused with a clear message when the run
is headless (`TF_BUILD`, `CI`, `-NonInteractive`, no console).

## 5. The token cache (DeviceCode)

File: `<BaseFolder>\State\auth\token-cache.json` (override with `-TokenCachePath` / `IMPACTIQ_TOKEN_CACHE_PATH`).
It is **never** inside a run folder, so it is never part of the outputs; the pipeline publishes `State\auth` only when
an AES key is set (a DPAPI blob is useless on another machine anyway).

Payload: `{ schemaVersion, authority, tenantId, clientId, environment, refreshToken, account, savedUtc }`, encrypted:

| Condition | Format | Readable by |
|---|---|---|
| `-TokenCacheKey` / `IMPACTIQ_TOKEN_CACHE_KEY` set | `{ "format": "aes256", "data": "<base64 IV+ciphertext>" }` - AES-256-CBC, key = SHA-256(key string), random 16-byte IV | anyone with the key, on any machine -> hosted agents |
| no key, Windows | `{ "format": "dpapi-user", "data": "..." }` - `ProtectedData.Protect(..., CurrentUser)` | the same Windows user on the same machine (domain accounts: also other domain-joined machines) -> self-hosted agents |
| no key, other OS | not persisted; one Warn per run | - |

`Restore-IQTokenCache` returns `$null` (and the tool starts a fresh device-code flow) when the file is missing,
corrupt, encrypted with a different key, or was created for another authority / client id / environment. A different
tenant id is tolerated (refresh tokens are not tenant-bound).

Lifetime: refresh tokens live 90 days of **inactivity** (sliding: every use issues a new one, which the tool saves) and
are revoked by a password change/reset, admin "revoke sessions", Conditional Access sign-in frequency, or a CA
**Authentication flows** policy that blocks device code (sessions started by device code stay "protocol-tracked", so a
later policy change kills existing caches too). A pipeline that runs at least weekly keeps the token alive indefinitely.

Rotating the key: change `IMPACTIQ_TOKEN_CACHE_KEY`, delete the old `impactiq-state` artifact's cache (or just let the
next run fail to decrypt it), run once by hand to sign in again.

## 6. Long runs

`Get-IQToken` refreshes proactively when fewer than **5 minutes** remain on the cached access token. Every Tabular
Editor XMLA process gets a fresh token in its connection string immediately before it starts. The v2 "log in again
every 55 minutes" timer is gone. `AccessToken` mode cannot refresh - use it only for short runs.

## 7. Things that break user auth (and what the log says)

| Symptom | Cause | Fix |
|---|---|---|
| `AADSTS50076` / `50079` / `53003` on Credential | MFA or Conditional Access required | DeviceCode (mode B/C) or an exempt account |
| `AADSTS530036` on device-code refresh | CA "Authentication flows" policy now blocks device code, refresh token revoked | ask the tenant admin to exclude the service account from that policy; re-bootstrap |
| `AADSTS50020` / `AADSTS90072` | account is a guest / wrong tenant | set `IMPACTIQ_TENANT_ID` |
| `AADSTS70016` / `expired_token` after 15 minutes | nobody completed the device code in time | run again while watching the log / webhook |
| `authorization_declined` | the human cancelled | run again |
| `invalid_grant` on a cached refresh token | password changed, sessions revoked, 90 idle days, different client id | the tool automatically starts a new device-code flow; complete it |
| `Failed to populate environments in settings` | agent cannot reach `api.powerbi.com` (needed even for USGov by the PowerBI module) | allow the host, or use DeviceCode/Credential (pure HTTP, no discovery call) |
| sign-in from a hosted agent rejected by a **named location** policy | hosted IPs are generic Azure ranges | self-hosted agent inside the trusted network, or a CA exclusion |
| `401` on every API call right after sign-in | token audience is wrong (e.g. commercial resource against `api.powerbigov.us`) | set `-Environment USGov`; do not use `-AuthorityOverride` to `login.microsoftonline.us` for GCC moderate |
| `403` on a specific workspace/dataset | the account lacks the role/permission | see Data-Coverage.md for the permission per sheet |

## 8. Security notes

* Never put secrets on the command line - `IMPACTIQ_*` environment variables are the documented path; the pipeline
  maps secret variables to them with `env:`.
* On Azure DevOps every minted access/refresh token is registered with `##vso[task.setsecret]` so the agent masks it in
  all log output. Log lines are additionally redacted (`Password=...`, `Bearer ...`, `access_token=...`).
* The manifest stores the mode and the account (`"auth": "DeviceCode (cached) as user@agency.gov"`), never tokens or
  passwords; `options` in the manifest drops any key matching password/secret/credential/token.
* The device-code message appears in the pipeline log (readable by everyone with pipeline read access). Whoever
  completes it first becomes the cached identity - integrity, not confidentiality, risk. Prefer the webhook to a
  private Teams channel and complete the code promptly.
* `IMPACTIQ_TOKEN_CACHE_KEY` + the `impactiq-state` artifact together are a long-lived credential for that user.
  Restrict who can download artifacts, and use an account with **no more than Viewer/Contributor** on the workspaces.
* Everything here is officially "not recommended" by Microsoft for automation - the supported answer is a service
  principal, which Power BI GCC supports. This edition exists for tenants where that is not (yet) possible.

## 9. AADSTS quick reference

`50076` MFA required (interactive) - `50079` MFA enrolment required - `53003` blocked by Conditional Access -
`65001` consent required - `50074` strong auth required - `530036` blocked by CA authentication-flows policy -
`50126` bad username/password - `50034` user not found - `50053` locked out - `50055` password expired -
`50056` no password - `50057` account disabled - `50059` / `90002` tenant not found - `50155` device auth required -
`70016` device code expired/pending - `700016` client id not found in tenant (wrong `-ClientId` for a sovereign cloud).
