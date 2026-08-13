# Security policy

## Supported versions

Security fixes are provided for the latest published release only. The theme
and Companion are deliberately constrained to Thunderbird 153.x because the
Companion uses privileged Thunderbird interfaces that can change between major
versions.

## Reporting a vulnerability

Please do not disclose a suspected vulnerability in a public issue. Use the
repository's **Security** tab and select **Report a vulnerability**:

https://github.com/UbhiTS/outlook-style-for-thunderbird/security/advisories/new

Include the affected version, Thunderbird version and operating system,
reproduction steps, impact, and any proof of concept. Avoid including real
message, contact, or calendar data. If private vulnerability reporting is not
available, contact the maintainer through their GitHub profile first and share
only enough information to arrange a private channel.

The maintainers aim to acknowledge reports within three business days and
provide an initial severity assessment within seven business days. Resolution
timing depends on severity and the availability of a safe Thunderbird-compatible
fix.

## Security model

- The static theme has no scripts, host permissions, or WebExtension
  permissions.
- The Companion has no WebExtension, host, optional, or externally-connectable
  permissions. It uses a packaged Experiment API to integrate with local
  Thunderbird windows.
- Neither package contains telemetry, remote code, network clients, or remotely
  loaded CSS and assets.
- Build artifacts are produced deterministically and published with SHA-256
  checksums. Users should verify checksums when installing outside an add-on
  store.

The Companion's elevated-access warning is expected because Experiment APIs run
with Thunderbird chrome privileges. A request to suppress that warning or add a
remote update path should be treated as a security-sensitive change.
