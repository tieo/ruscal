# Code Signing — current status

`ruscal` ships **unsigned**. Microsoft Defender's cloud ML classifier has
been observed flagging the binary as `Trojan:Win32/Bearfoos.A!ml` — a
heuristic-only label assigned to "looks like a dropper" patterns. The fix
is a real Authenticode signature anchored to a publisher identity.

## Free signing paths surveyed (2026)

| programme | status | usable now? |
|---|---|---|
| SignPath Foundation OSS | declined — project not popular enough yet | no, reapply once usage grows |
| OSSign (ossign.org) | applications suspended due to backlog | no, watch for queue reopening |
| Certum free OSS cert | discontinued in 2016 | no |
| Sectigo / Comodo free OSS cert | discontinued | no |
| Microsoft Trusted Signing | ~$10/month Azure | paid, not free |
| Certum paid OSS cert | ~€69 first year, €29/yr renewal | paid, not free |

No free Authenticode signing path is currently open. Until that changes
`ruscal` is distributed unsigned.

## What this means for users

When you download `ruscal.exe` from the GitHub release page (or pull it
through WinGet), Windows Defender may quarantine it as Bearfoos. If it
does:

1. Open Windows Security → Virus & threat protection → Protection history.
2. Find the entry for `ruscal.exe`.
3. Choose **Actions → Restore**, then **Allow on device**.
4. Optional: submit the SHA-256 hash to
   <https://www.microsoft.com/en-us/wdsi/filesubmission> as
   "Software developer" → "Incorrect detection" so Microsoft adds it to
   their cleared-hashes list.

The build does not contain a real trojan. The source code is public at
<https://github.com/tieo/ruscal>; the behaviours capa identifies (registry
write under `HKCU\…\Run`, Start-menu shortcut creation, MAPI calls,
clipboard write, TLS / OAuth crypto) are all needed by the app and
explained in commit messages.

## When signing becomes possible again

The dev-side actions when a free signing path opens:

1. Re-add a SignPath (or equivalent) signing job to `.github/workflows/release.yml`.
2. Add `SIGNPATH_API_TOKEN` (secret) and `SIGNPATH_ORG_ID` (variable) — or
   the equivalent for whichever programme accepts the project.
3. Tag a release; first 2–3 signed builds may still need WDSI submissions
   while Defender's model accumulates reputation against the certificate.

## Repo secrets the existing workflow expects

`release.yml` is wired for an auto-WinGet-publish step
(`vedantmgoyal2009/winget-releaser@v2`). It requires a single repo secret:

* `WINGET_TOKEN` — a classic GitHub Personal Access Token, `public_repo`
  scope (no other scopes needed). The action uses it to push a branch to
  `tieo/winget-pkgs` (the existing fork of `microsoft/winget-pkgs`) and
  open the manifest-update PR. Create at
  <https://github.com/settings/tokens/new>, store under
  Repo Settings → Secrets and variables → Actions → New repository secret.

Without `WINGET_TOKEN` set, only the `winget` job in the workflow fails;
the release itself still publishes.
