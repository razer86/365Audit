# Security Policy

## Supported Versions

Only the latest release on `main` is actively maintained. Security fixes are not backported to older versions.

## Reporting a Vulnerability

**Do not open a public GitHub issue for security vulnerabilities.**

Please report security issues by emailing the maintainer directly. You can find contact details via the GitHub profile. Include:

- A description of the vulnerability and its potential impact
- Steps to reproduce or proof-of-concept (if safe to share)
- Any suggested miigation

You should receive an acknowledgement within 5 business days. If you do not, follow up via a GitHub DM.

## Scope

This toolkit reads from Microsoft 365 tenants using delegated or application permissions. Relevant security concerns include:

- Credential or token exposure in script output, logs, or generated reports
- Overly broad permission requests beyond what auditing requires
- Injection risks in report generation (e.g., tenant data rendered as HTML)

## Out of Scope

- Vulnerabilities in Microsoft 365 itself or its APIs
- Issues requiring physical access to the machine running the scripts
