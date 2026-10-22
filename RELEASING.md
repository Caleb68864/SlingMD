# Releasing SlingMD

Releases are built on Windows with Visual Studio (Office/VSTO workload). The `dotnet` CLI can't
build this project, and the CLI `msbuild -t:Publish` target doesn't regenerate the ClickOnce
payload for VSTO projects, so the steps below are the only supported path.

1. **Set the version.** Update `ApplicationVersion` in `SlingMD.Outlook/SlingMD.Outlook.csproj`
   and `AssemblyFileVersion` in `SlingMD.Outlook/Properties/AssemblyInfo.cs`. Move the
   `CHANGELOG.md` entry from *Unreleased* to the version and date.
2. **Build and test.** Build Debug and run the tests (commands in `CLAUDE.md`), then try the main
   flows in Outlook.
3. **Publish from Visual Studio.** *Build → Publish SlingMD.Outlook*. This writes
   `SlingMD.Outlook/publish/Application Files/SlingMD.Outlook_<version>/`.
4. **Package.** Run `.\package-release.ps1`. It rebuilds Release and refuses to package if the
   published DLL doesn't match the fresh build, then writes a versioned ZIP to `Releases/`.
5. **Tag and upload.** Tag the commit `v<version>` and attach the ZIP to a GitHub release.

The old `build-and-publish*.ps1` scripts are kept in `archive/` for reference only. They use the
`dotnet` CLI or CLI publish and can produce a stale payload, so don't use them for releases.
