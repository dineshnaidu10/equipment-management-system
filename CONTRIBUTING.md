# Contributing

Thanks for your interest in improving EMS!

## Reporting bugs
Open an issue and include:
- Operating system and Python version (`python --version`)
- Steps to reproduce
- What you expected vs. what happened
- Any traceback shown in the terminal

**Never attach real databases, config files or attachments containing company data.** Use the sample data instead.

## Suggesting features
Open an issue describing the problem you are trying to solve before writing code. Because the project is deliberately a **single `.py` file** with no server and minimal dependencies, features that need extra infrastructure may be declined.

## Pull requests
1. Fork the repo and create a branch: `git checkout -b feature/short-description`
2. Keep the project single-file and avoid new required dependencies (optional ones are fine if the app degrades gracefully).
3. Test manually against `sample_data/ems_config.xlsx` on a fresh database.
4. Update `README.md` and `CHANGELOG.md` if behaviour changes.
5. Open the PR with a clear description of what and why.
