# ruscal

Stops you from forgetting outlook calendar events by clearly showing them in *your* calendar.

Requires Outlook Classic; the new Outlook does not expose MAPI.

## Install

```
winget install tieo.ruscal
```

Updates: `winget upgrade tieo.ruscal`. There is no in-app updater.

## Heads up — Defender false positive

`ruscal` ships **unsigned** (see [docs/SIGNING.md](docs/SIGNING.md) for why).
Microsoft Defender may quarantine it as `Trojan:Win32/Bearfoos.A!ml` —
a false positive. Restore via Windows Security → Protection history →
Actions → Restore, then "Allow on device". The source is public; nothing
in it is a trojan.

---

Licensed under [MIT](LICENSE). See [TERMS.md](TERMS.md) for
use-at-your-own-risk and [PRIVACY.md](PRIVACY.md) for what data ruscal
touches. **You are responsible for ensuring your use complies with your
organisation's policies, connected services' terms, and applicable law.**
