---
name: outlook-addin-load-diagnostics
description: Diagnose which manifest Outlook is using and why a sideloaded add-in fails to load even when the local dev server is reachable in a browser.
---

# Outlook Add-in Load Diagnostics

Use this skill when the user reports that the add-in fails to open in Outlook but the local app URL works in a browser.

## When To Use

- User asks which manifest Outlook is actually using.
- User sees a generic Outlook add-in load error.
- Local dev server works in browser but Outlook host fails.
- Need host-side diagnostics from Office runtime logs.

## Workflow

1. Confirm active development registration:
   - Run: npx office-addin-dev-settings registered
   - Match add-in ID to manifest path.
   - If unclear, clear and re-register:
     - npx office-addin-dev-settings clear manifest.json
     - npx office-addin-dev-settings register manifest.json
     - npx office-addin-dev-settings registered

2. Verify manifest runtime endpoints:
   - Read manifest and check runtime code.page and icon URLs.
   - Confirm they match the expected local or hosted endpoint.

3. Enable host diagnostics:
   - Enable runtime log:
     - npx office-addin-dev-settings runtime-log --enable
   - Log path is typically:
     - C:\Users\<User>\AppData\Local\Temp\OfficeAddins.log.txt
   - Enable debugging:
     - npx office-addin-dev-settings debugging manifest.json --enable --debug-method direct --open-dev-tools

4. Validate localhost access from New Outlook app container:
   - Check loopback exemptions:
     - CheckNetIsolation LoopbackExempt -s
   - Get package family name:
     - Get-AppxPackage _Outlook_ | Select-Object Name,PackageFamilyName
   - Add exemption if missing:
     - CheckNetIsolation LoopbackExempt -a -n="<PackageFamilyName>"
   - Verify again with CheckNetIsolation LoopbackExempt -s

5. Reproduce and collect errors:
   - Restart Outlook completely.
   - Start sideload flow from repo.
   - Open add-in and capture newest OfficeAddins.log entries.
   - Filter for add-in ID, localhost URL, and error markers.

## Interpretation Guide

- registered points to expected manifest, but add-in still fails:
  - Most likely host runtime/network policy issue, cache issue, or auth/redirect mismatch.

- Browser can open localhost, Outlook cannot:
  - Strong signal of missing loopback exemption for packaged Outlook host.

- Runtime log has parse or URL errors:
  - Manifest content issue or malformed URL value.

- Runtime log has auth/redirect errors:
  - Azure app redirect URI mismatch with active runtime page URL.

## Guardrails

- Do not change manifest ID unless explicitly asked.
- Do not change production permissions or app IDs as part of diagnosis.
- Keep fixes minimal and reversible.
- Prefer diagnostic commands first, code changes second.

## Output To User

Return:

- The manifest path Outlook is currently registered to use.
- The most likely root cause category.
- Exact commands run and key findings.
- Concrete next actions to verify the fix.
