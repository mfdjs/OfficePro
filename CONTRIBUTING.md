# Contributing to OfficePro

Thanks for your interest in improving OfficePro! This document explains how to report issues, set up a development environment, and submit changes.

## Code of Conduct

Be respectful and constructive. We welcome contributors of all backgrounds and skill levels. Harassment or abuse will not be tolerated.

## Reporting bugs

Before opening an issue, please search existing issues to avoid duplicates. When you file a bug, include:

- What you did (clear steps to reproduce)
- What you expected to happen
- What actually happened (screenshots or the full error text are very helpful)
- The file type involved (e.g. `.docx`, `.xlsx`, `.pdf`)
- Your OS and OfficePro version
- Any console / DevTools error output

## Suggesting features

Describe the problem you want to solve, not just the proposed solution. Explain your use case and why it would be useful.

## Development setup

```bash
git clone https://github.com/<your-username>/OfficePro.git
cd OfficePro
npm install
npm start
```

See the [README](README.md) for OCR data setup and build instructions.

## Coding guidelines

- OfficePro uses **vanilla HTML/CSS/JavaScript** on top of Electron; match the existing style rather than introducing a new framework unless it has been discussed in an issue.
- Keep code self-explanatory. Do **not** commit `console.log`, debugging code, or leftover scratch comments.
- Use clear, descriptive names; keep functions small and focused.
- Preserve a consistent indentation style (spaces) and end-of-line behavior.
- Save all files as **UTF-8** and avoid introducing non-ASCII characters where plain ASCII is sufficient.
- Handle errors explicitly: parsing and file operations should fail gracefully with a clear message, never crash or show garbled output.
- Do not translate user document content when adding UI text; only the interface is localized.

## Adding or improving translations

- UI translations live in [`locales/`](locales) and are loaded by [`i18n.js`](i18n.js).
- The translation key is the original Chinese string; keys must match exactly, including punctuation.
- Do not machine-translate blindly — review wording for naturalness. Right-to-left languages are supported automatically.
- After editing, verify the language switch in the running app.

## Testing

Please exercise your changes in the real app:

1. Open the relevant file types and confirm they render correctly.
2. Make sure the buttons/commands you touch actually perform their action (no empty "done" toasts).
3. Check the DevTools console for errors.
4. If a test suite exists, run it and ensure nothing regressed.

## Commit messages

Use concise, descriptive messages (for example `fix: handle encrypted docx with clear message`). Prefer small, focused commits.

## Pull request checklist

- [ ] My code follows the coding guidelines above
- [ ] I removed debug logging and scratch comments
- [ ] I tested the change in the running app, including the affected file types
- [ ] Buttons/commands I changed really work
- [ ] I did not break unrelated modules
- [ ] I updated documentation where needed

## Workflow

1. Fork the repository and create a branch from `main`.
2. Make and test your changes.
3. Push your branch and open a pull request, describing what and why.
4. A maintainer will review it and may request changes.

Thank you for contributing!
