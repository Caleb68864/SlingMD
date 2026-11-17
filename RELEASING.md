# Releasing SlingMD

Releases are built on Windows with Visual Studio (Office/VSTO workload). The `dotnet` CLI can't
build this project, and the CLI `msbuild -t:Publish` target doesn't regenerate the ClickOnce
payload for VSTO projects, so the steps below are the only supported path.

1. **Set the version.** Pick the release number and set it in both places *before* publishing:
   `ApplicationVersion` in `SlingMD.Outlook/SlingMD.Outlook.csproj` and `AssemblyFileVersion` in
   `SlingMD.Outlook/Properties/AssemblyInfo.cs`. The file version is compiled into the DLL, so
   changing it after publishing means publishing and packaging again.
2. **Build and test.** Build Debug and run both test suites (commands in `CLAUDE.md`):
   `SlingMD.Tests` with VSTest and `dotnet test SlingMD.Core.Tests`. Then try the main flows in
   Outlook.
3. **Publish from Visual Studio.** *Build → Publish SlingMD.Outlook*. This writes
   `SlingMD.Outlook/publish/Application Files/SlingMD.Outlook_<version>/`.
   `AutoIncrementApplicationRevision` is on, so Visual Studio can bump the last number of
   `ApplicationVersion` in the csproj when it publishes. Afterwards, check the newest
   `SlingMD.Outlook_<version>` folder and the csproj: that is the version you are releasing. If it
   no longer matches `AssemblyFileVersion`, either accept the mismatch (installs and updates go by
   `ApplicationVersion`) or fix `AssemblyFileVersion` and publish again.
4. **Package.** Run `.\package-release.ps1`. It rebuilds Release and refuses to package if the
   published DLL doesn't match the fresh build, then writes a versioned ZIP to `Releases/`.
   The ZIP also gets `SlingMD.cer` (taken from the published manifest) and `Trust-SlingMD.cmd`/`.ps1`
   from `install/`, for computers that refuse the self-signed certificate.
5. **Date the changelog.** Rename `## [Unreleased]` in `CHANGELOG.md` to
   `## [<version>] - <yyyy-MM-dd>`. Don't add a new empty *Unreleased* section; the next change adds
   it.
6. **Commit.** Commit the csproj, `AssemblyInfo.cs`, `CHANGELOG.md` and the new ZIP in `Releases/`
   as `Release <version>`, then push.
7. **Tag.** Create an annotated tag on that commit and push it:
   `git tag -a v<version> -m "SlingMD <version>"` and `git push origin v<version>`.
8. **Create the GitHub release** with the ZIP attached, titled `SlingMD <version>`:
   `gh release create v<version> Releases/SlingMD.Outlook_<version_with_underscores>.zip --title "SlingMD <version>" --notes-file <notes> --verify-tag --latest`.
   The notes start with the *Installing* section from the previous release
   (`gh release view v<previous> --json body --jq .body`) with the ZIP name updated, followed by
   this version's `CHANGELOG.md` section. The *Installing* section carries the signing
   certificate's thumbprint so people can check what `Trust-SlingMD.cmd` shows them. Before
   publishing, check that the `SlingMD.cer` inside the ZIP has that thumbprint. The current
   certificate is `CN=SlingMD Development`, `736CF7BD97289EC0E7B925A4364C54B1822892F2`, valid
   until 2031-03-25.

**Keep the signing key.** `SlingMD.Outlook_TemporaryKey.pfx` is git-ignored and lives only on the
release machine. Back it up. Installed copies only accept updates signed with the same certificate,
so a new key means everyone has to uninstall and reinstall, and trust the new certificate again.

**Committing from PowerShell:** pass a multi-line message with `git commit -m @'...'@` (the closing
`'@` at the start of its line). `git commit -F - @'...'@` treats the text as a file path and fails.

The old `build-and-publish*.ps1` scripts are kept in `archive/` for reference only. They use the
`dotnet` CLI or CLI publish and can produce a stale payload, so don't use them for releases.
