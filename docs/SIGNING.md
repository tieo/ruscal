# Code Signing & Defender False-Positive Workflow

`ruscal` was flagged as `Trojan:Win32/Bearfoos.A!ml` by Defender's cloud ML
classifier on 2026-06-21. `!ml` means the classifier matched a behavioural
*pattern*, not a known signature. The triggering pattern was the combination
of:

1. A GUI exe living in `%LOCALAPPDATA%`.
2. An `HKCU\Software\Microsoft\Windows\CurrentVersion\Run\ruscal` autostart
   key.
3. A Start-menu `.lnk` pointing into `%LOCALAPPDATA%`.
4. Unsigned PE.

(1), (2), (3) are intrinsic to the product — that's how ruscal installs and
starts. The only mitigation is to **anchor the binary's identity** with an
Authenticode signature so Defender stops evaluating it on heuristic alone.

A separate pre-existing trigger — the GUI exe spawning `powershell.exe` for
shortcut creation, clipboard, and process termination — has been removed in
commit `af2a2b2`. All those operations now use native Win32 calls.

## Path 1 — Sign every release via SignPath Foundation (free for OSS)

The `release.yml` workflow already invokes
`signpath/github-action-submit-signing-request@v1`. To turn it on you need a
one-time SignPath Foundation enrolment:

### 1. Apply for the SignPath Foundation OSS plan

https://signpath.org/apply

Fill in the form:

* **Project name:** `ruscal`
* **Project URL:** `https://github.com/tieo/ruscal`
* **License:** GPL-3.0 (or whatever `LICENSE` says — must match).
* **Maintainer:** the GitHub account that owns the repo.

SignPath replies in a few business days with an organisation ID, a signing
policy slug, and an API token.

### 2. Configure the SignPath project

In the SignPath portal:

* **Project slug:** `ruscal` (matches the workflow).
* **Artifact configuration slug:** `ruscal-exe` (matches the workflow).
  Artifact type = "Single file", expected file = `ruscal.exe`.
* **Signing policy slug:** `release-signing` (matches the workflow).
  Policy = "Release", trusted submitter = your GitHub repo `tieo/ruscal`,
  branch/tag pattern = `refs/tags/v*`.

### 3. Add GitHub secrets / variables

In `https://github.com/tieo/ruscal/settings/secrets/actions`:

| Kind     | Name                  | Value                                   |
| -------- | --------------------- | --------------------------------------- |
| Secret   | `SIGNPATH_API_TOKEN`  | The API token issued by SignPath.       |
| Variable | `SIGNPATH_ORG_ID`     | The organisation ID issued by SignPath. |

`SIGNPATH_ORG_ID` is non-sensitive (it's just a UUID) — set it as a
*variable* not a *secret* so it shows up in workflow logs for debugging.

### 4. Cut a release

```bash
git tag v1.2.0
git push origin v1.2.0
```

The workflow builds the binary, uploads it as an unsigned artifact,
SignPath fetches it, signs it, and the signed binary is published to the
GitHub release. **Do not run any locally-built dev binary** until SignPath
has signed at least one release — Defender will quarantine it again.

### 5. First few releases: re-submit to WDSI per release

Even with a valid signature, Defender's ML model needs a handful of
"this signature is benign" data points before it stops flagging. For the
first 2–3 signed releases:

1. Download the signed `ruscal.exe` from the GitHub release.
2. Submit to <https://www.microsoft.com/en-us/wdsi/filesubmission> as
   "Software developer" → "Incorrect detection".
3. Mention previous detection name (`Trojan:Win32/Bearfoos.A!ml`) in the
   notes and that the binary is now Authenticode-signed via SignPath.
4. MS clears the hash within 24–72h and re-trains the model.

Once reputation builds, future releases auto-clear without further action.

## Path 2 — Stay unsigned forever

If SignPath signup fails for any reason, the fallback is per-release
WDSI submission for every new `ruscal.exe` hash. Free, but tedious and
no reputation accumulates because there's no publisher identity to
attach reputation to.

## SHA256 of the quarantined binary (for WDSI submission)

The detection on 2026-06-21 21:03 was on
`%LOCALAPPDATA%\ruscal\ruscal.exe` produced from commit `0957311`. To
get the exact hash:

```powershell
Get-FileHash -Algorithm SHA256 -Path "$env:LOCALAPPDATA\ruscal\ruscal.exe"
```

(If Defender has already quarantined the file, restore via
Windows Security → Protection history before hashing — or take the SHA256
from the latest GitHub release asset, which has identical bytes.)
