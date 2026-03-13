# Changelog

All notable changes to SlingMD are documented in this file.

## [1.0.0.124] - 2026-03-13

### Fixed
- Corrupt or malformed `ObsidianSettings.json` no longer crashes the add-in on startup; safe defaults are loaded instead ([#6](https://github.com/Caleb68864/SlingMD/issues/6))
- Settings now persist correctly across Outlook restarts and Sling operations ([#6](https://github.com/Caleb68864/SlingMD/issues/6))
- Fatal export errors no longer fall through into contact creation or Obsidian launch
- Canceling the task-options dialog no longer disables task creation for the rest of the Outlook session
- Thread summary notes (0-file) DataviewJS query now correctly scopes to the thread folder instead of producing a `SyntaxError: Invalid or unexpected token`
- Thread discovery now parses both second-precision (`yyyy-MM-dd HH:mm:ss`) and legacy minute-precision (`yyyy-MM-dd HH:mm`) date formats
- Frontmatter is now YAML-safe when email metadata contains double quotes, backslashes, or embedded newlines
- Attachment links now resolve correctly for same-folder, per-note-subfolder, and centralized attachment storage modes
- Exporting into a fresh vault with a missing inbox folder no longer throws during duplicate detection or cache initialization
- Contact notes now use `## Communication History` heading with a working DataviewJS query for email history ([#4](https://github.com/Caleb68864/SlingMD/issues/4))

### Added
- Customizable markdown templates for email notes, contact notes, task lines, and thread summaries via Settings ([#8](https://github.com/Caleb68864/SlingMD/issues/8))
- Regression test coverage for corrupt-settings fallback, task-state reset, missing inbox handling, thread date compatibility, frontmatter escaping, and attachment-link generation
- VSTO build/test prerequisite documentation in README

## [1.0.0.121] - 2025-12-15

### Fixed
- Outlook settings persistence improvements
- Contact history heading alignment

## [1.0.0.44] - 2025-03-15

### Added
- Automatic email thread detection and organization
- Thread summary pages with timeline views
- Configurable subject cleanup patterns
- Thread folder creation for related emails
- Participant tracking in thread summaries
- Dataview integration for thread visualization

### Improved
- Email relationship detection
- Thread navigation with bidirectional links

## [1.0.0.14] - 2025-02-01

### Added
- Follow-up task creation in Obsidian notes
- Follow-up task creation in Outlook
- Configurable due dates and reminder times
- Task options dialog for custom timing

## [1.0.0.8] - 2025-01-15

### Added
- Initial release
- Email to Obsidian note conversion
- Email metadata preservation
- Obsidian vault configuration
- Launch delay settings
