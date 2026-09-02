# Security Policy

Thanks for helping keep DocConvert and its users safe.

## Supported Versions

The latest released version on PyPI (`docconvert-local`) is the only line that
receives security fixes. Older versions are not patched; please upgrade before
filing a report.

| Version | Supported |
| ------- | --------- |
| `2.0.x` (latest) | ✅ Active |
| `< 2.0.0` | ❌ End of life |

## Reporting a Vulnerability

Please **do not** file security issues as public GitHub Issues or Discussions.
Use one of the private channels below instead:

- **Email**: `leop017@users.noreply.github.com` (GitHub-noreply alias — please
  use the address shown on the maintainer's commit author line if a different
  one is configured)
- **GitHub private vulnerability report**: click "Report a vulnerability" on
  the [Security Advisories](https://github.com/leop017/DocConvert/security/advisories)
  tab — recommended, since it stays on GitHub and keeps the conversation
  auditable.

### What to include

A useful report contains:

1. A clear description of the vulnerability and its impact.
2. Reproduction steps or a minimal PoC (script, sample document, command line).
3. The affected version (`pip show docconvert-local`).
4. Your environment: Python version, OS, install path.

## Response Timeline

This is a single-maintainer project, so response windows are best-effort:

| Step | Target |
| ---- | ------ |
| Initial acknowledgement | within **7 days** |
| Triage + severity assessment | within **14 days** |
| Patch release for `Critical` / `High` | within **30 days** |
| Patch release for `Medium` / `Low` | next regular release |

If you need to escalate (no reply after 14 days), leave a public comment on
the related advisory asking for status — that does not leak details.

## Coordinated Disclosure

We follow a **90-day responsible disclosure** window. After a fix ships:

- The advisory is published with full details.
- A `GHSA-xxxx-xxxx-xxxx` ID is assigned.
- A `CHANGELOG.md` entry is added under the next release.

We will credit reporters in the advisory (unless you ask to stay anonymous)
and in the release notes.

## Scope

In scope for this project:

- `docconvert-local` Python package (all submodules under `docconvert/`)
- The Tk GUI shipped with the package
- PyInstaller-built executables attached to GitHub Releases
- CI / release workflows under `.github/workflows/`

Out of scope:

- Third-party dependencies (`python-docx`, `openpyxl`, `mammoth`, …) — please
  report those upstream.
- Issues that require the user to open an attacker-controlled file from an
  untrusted source **and** to execute it in a way that escapes the
  converter's normal usage patterns.

## Out-of-Band Updates

For `Critical` issues (e.g. arbitrary code execution from a malicious
workbook), we may cut an out-of-band `2.0.x` patch without waiting for the
regular release cadence. A `YANKED` record in `CHANGELOG.md` will mark any
release that turns out to be unsafe.