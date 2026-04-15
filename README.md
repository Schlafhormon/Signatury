# Signatury

Signatury is a PowerShell-based Outlook signature deployment solution for Windows Active Directory environments. It uses existing Word `.docx` templates as the master source, resolves user data from Active Directory, generates Outlook-compatible signature files, and sets the default signature for Classic Outlook on Windows.

The project is designed for local testing first, but the architecture is ready for central rollout through GPO logon scripts, scheduled tasks, or Intune script execution in user context.

## Project Status

Signatury is a solo-maintained project. The code is published publicly so anyone can use it, study it, and fork it under the terms of the [MIT License](LICENSE), but active development is maintained by the repository owner.

## Highlights

- Reuses existing Word signature templates such as `DK_Standard.docx`
- Replaces placeholders reliably through Microsoft Word COM automation
- Preserves embedded logos and images in exported HTML signatures
- Reads user attributes directly from on-prem Active Directory
- Generates `.htm`, `.txt`, and optional `.rtf`
- Sets default signatures for new emails and replies/forwards in Classic Outlook
- Uses a local cache and fingerprint-based idempotency
- Works in user context without admin rights during normal operation
- Requires no cloud, Graph, Entra ID, or third-party signature tooling

## Architecture

Signatury uses Word as the rendering engine because the template source of truth is a `.docx` file. This allows the solution to keep layout fidelity, tables, fonts, and embedded images while still doing placeholder replacement robustly through Word's own search and replace engine.

The code is split into distinct responsibilities:

- Configuration loading
- Active Directory data lookup
- Template caching
- Word-based template processing
- Signature export
- Outlook default-signature registry updates
- Logging and state tracking

The main implementation lives in [Scripts/OutlookSignature.Core.psm1](Scripts/OutlookSignature.Core.psm1).

## Repository Structure

```text
Templates/
  DK_Standard.docx
Config/
  signature-config.json
  signature-config.enterprise.example.json
Scripts/
  OutlookSignature.Core.psm1
  Deploy-OutlookSignature.ps1
  Invoke-OutlookSignatureRollout.ps1
Tests/
  Invoke-OutlookSignatureSelfTest.ps1
Docs/
  Administrator-Guide.md
  README_GER.md
Logs/
Temp/
  Cache/
  State/
  Work/
```

## Supported Placeholder Mapping

The following placeholders are implemented by default:

- `$CurrentUserGivenName$` -> `givenName`
- `$CurrentUserSurname$` -> `sn`
- `$CurrentUserTitle$` -> `title`
- `$CurrentUserTelephone$` -> `telephoneNumber`
- `$CurrentUserMail$` -> `mail`

Additional placeholders can be added through the configuration file without redesigning the core architecture.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1
- Microsoft Word installed locally
- Classic Outlook for Windows
- Active Directory domain user context for production use

## Quick Start

1. Place your template in `Templates\DK_Standard.docx`.
2. Review [Config/signature-config.json](Config/signature-config.json).
3. Open Windows PowerShell in the repository root.
4. Run:

```powershell
.\Scripts\Deploy-OutlookSignature.ps1
```

Optional:

- Skip Outlook default-signature updates:

```powershell
.\Scripts\Deploy-OutlookSignature.ps1 -SkipDefaultSignatureUpdate
```

- Force a full rebuild:

```powershell
.\Scripts\Deploy-OutlookSignature.ps1 -ForceRefresh
```

## Rollout

For enterprise rollout, switch the local paths in the configuration to central UNC paths and place cache/log/temp/state folders under `%LOCALAPPDATA%\Signatury`.

Example GPO logon command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "\\fileserver\Signatury\Scripts\Invoke-OutlookSignatureRollout.ps1" -ConfigPath "\\fileserver\Signatury\Config\signature-config.json"
```

An enterprise-ready example configuration is included in [Config/signature-config.enterprise.example.json](Config/signature-config.enterprise.example.json).

## Logging and Idempotency

Each run writes a per-user log file and stores a state file with a fingerprint derived from:

- template hash
- resolved variable values
- signature options

If nothing relevant changed, Signatury keeps the existing generated signature files and only refreshes Outlook registry defaults.

## Documentation

- German README: [Docs/README_GER.md](Docs/README_GER.md)
- Administrator guide: [Docs/Administrator-Guide.md](Docs/Administrator-Guide.md)
- Publishing guide: [Docs/PUBLISHING.md](Docs/PUBLISHING.md)
- Local test config: [Config/signature-config.json](Config/signature-config.json)
- Enterprise example config: [Config/signature-config.enterprise.example.json](Config/signature-config.enterprise.example.json)
- Templates guide: [Templates/README.md](Templates/README.md)
- Contributing: [CONTRIBUTING.md](CONTRIBUTING.md)
- Security policy: [SECURITY.md](SECURITY.md)

## Notes

- The project currently targets Classic Outlook for Windows.
- The implementation intentionally avoids external signature products and online dependencies.
- RTF export is supported on a best-effort basis through Word and is logged clearly if it fails for a specific template.
- The included `DK_Standard.docx` is a neutral sample template for testing and can be replaced with your own corporate template.
- Public use is explicitly intended through the included [MIT License](LICENSE).
