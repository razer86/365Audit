# Contributing to 365 Audit Toolkit

Thanks for your interest in contributing. This is a PowerShell-based MSP audit toolkit — contributions that improve coverage, accuracy, or usability are welcome.

## Getting Started

1. Fork the repo and clone locally
2. Create a branch: `git checkout -b feat/your-feature`
3. Make your changes and test against a real Microsoft 365 tenant
4. Submit a pull request against `main`

## AI-Assisted Development

This codebase is developed with the assistance of AI tools (Claude by Anthropic). What that means for contributors:

- All AI-generated code is human-reviewed and tested against live Microsoft 365 tenants before being committed
- PRs are held to the same standard — do not submit AI-generated code that hasn't been validated in a real environment
- If you used AI assistance to write your contribution, say so in the PR description

## Guidelines

- **PowerShell 7.4+** — all scripts must be compatible
- **No breaking changes to config.psd1 structure** without a migration note
- Follow the existing parameter naming conventions (`-TenantId`, `-CertThumbprint`, etc.)
- Update `version.json` with a bumped version for any changed script
- Add an entry to `CHANGELOG.md` under `[Unreleased]`

## Reporting Bugs

Use the [Bug Report](.github/ISSUE_TEMPLATE/bug_report.md) issue template. Include your PowerShell version, module versions, and the full error output.

## Feature Requests

Use the [Feature Request](.github/ISSUE_TEMPLATE/feature_request.md) template. If the request relates to a specific Microsoft 365 workload (Entra, Exchange, etc.), mention which module it belongs to.
