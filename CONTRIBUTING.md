# Contributing

Thank you for improving Outlook Style for Thunderbird. Keep changes focused and
preserve Thunderbird's native mail, calendar, invitation, and security behavior
wherever possible.

## Development workflow

1. Discuss substantial UI behavior or privileged API changes before investing
   in a large patch.
2. Work from the latest `main` branch and keep unrelated changes separate.
3. Update both manifests together when changing the version or compatibility
   range, and add release notes to `README.md`.
4. Run the production gate before opening a pull request:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\release-gate.ps1
   ```

5. Describe the Thunderbird version, operating system, layout, theme mode, and
   manual scenarios tested. Include before/after screenshots for visible UI
   changes.

## Security and privacy boundaries

- Do not add WebExtension permissions, remote assets, telemetry, analytics,
  network clients, dynamic code evaluation, or HTML string injection without a
  separate threat-model review and explicit maintainer approval.
- Keep the Experiment API surface narrow. Schema changes and new privileged
  lifecycle hooks require security and shutdown/cleanup review.
- Never include real mail, contact, attendee, calendar, profile, or credential
  data in tests, logs, screenshots, or fixtures.
- Report suspected vulnerabilities through [SECURITY.md](SECURITY.md), not in a
  public issue.

## Compatibility

Internal Thunderbird selectors and APIs can change without notice. ATN requires
Experiment add-ons to declare `strict_max_version`, so test clean install,
upgrade, startup, shutdown, disable/enable, and affected workflows before
advancing that limit for each new Thunderbird major version.
