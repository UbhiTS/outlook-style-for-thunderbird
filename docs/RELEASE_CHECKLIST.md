# Production release checklist

Releases are intentionally manual. GitHub Actions builds every pushed commit but
does not create tags or GitHub Releases.

## Before merging

- Review the complete diff, including generated selectors and privileged API
  lifecycle changes.
- Keep the theme and Companion versions identical and update release notes.
- Confirm both manifests retain their stable Gecko IDs, least-privilege fields,
  and the tested Thunderbird minimum and maximum versions.
- Run the full gate from the repository root:

  ```powershell
  powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\release-gate.ps1
  ```

- Test install, upgrade, disable, enable, and uninstall paths in a disposable
  Thunderbird profile. Exercise light, dark, forced-colors, and reduced-motion
  modes where the changed feature is visible.
- For privileged behavior, test on every operating system claimed in the release
  notes. Never widen `strict_max_version` without a clean install and regression
  pass on that Thunderbird major version.

## Build the release candidate

1. Commit the approved source changes.
2. From that clean commit, run:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\release-gate.ps1 -RequireClean
   ```

3. Push the commit and wait for **Build add-ons** to pass.
4. Download the workflow artifact. Confirm it contains exactly the two XPI
   files, two matching ZIP files, and `SHA256SUMS.txt`.
5. Compare the downloaded checksums with the locally gated artifacts. A mismatch
   is a release blocker.

## Publish manually

- Create a signed tag from the exact reviewed commit.
- Publish immutable artifacts from the successful workflow run; do not rebuild
  them on another workstation.
- Include the supported Thunderbird range, elevated Companion access warning,
  privacy statement, checksum file, upgrade notes, and known limitations.
- If using an add-on store, complete its submission and review process before
  directing general users to the package.

## After publication

- Install the published artifacts into a fresh profile and repeat a focused
  smoke test.
- Verify the public SHA-256 values and download links.
- Monitor installation and startup failures. For a severe regression, withdraw
  the affected artifacts, document the last known-good version, and ship a new
  version rather than replacing files under an existing release.
