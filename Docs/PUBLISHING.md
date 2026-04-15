# Publishing Signatury on GitHub

## Before First Public Push

- Review `LICENSE` and confirm MIT matches your intended licensing.
- Verify that `Templates/DK_Standard.docx` is the neutral sample template and not an internal company document.
- Verify that no local logs, cache files, or temporary files are included.
- Review `Config/signature-config.enterprise.example.json` and adjust example paths if you want to show your preferred naming scheme.

## Recommended Initial Repository Files

- `README.md`
- `LICENSE`
- `CONTRIBUTING.md`
- `SECURITY.md`
- `CODE_OF_CONDUCT.md`
- `Docs/README_GER.md`

## Suggested First Release Steps

1. Create a new public GitHub repository named `Signatury`.
2. Initialize git locally if needed.
3. Commit the current repository contents.
4. Push to GitHub.
5. Add a short repository description.
6. Add topics such as `powershell`, `outlook`, `signature`, `active-directory`, `windows`, and `office`.
7. Create an initial release tag such as `v0.1.0`.

## Example Git Commands

```powershell
git init
git add .
git commit -m "Initial public release of Signatury"
git branch -M main
git remote add origin https://github.com/<your-user-or-org>/Signatury.git
git push -u origin main
```

## Recommended Repository Description

`PowerShell-based Outlook signature deployment for Windows Active Directory environments using Word DOCX templates.`
