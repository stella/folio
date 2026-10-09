# Security Policy

## Reporting a Vulnerability

**Do not open a public issue for security vulnerabilities.**

If you discover a security vulnerability in Folio, please report it
privately so we can address it before public disclosure:

- **Email:** [security@stll.app](mailto:security@stll.app)
- **GitHub Security Advisories:**
  [Report a vulnerability](https://github.com/stella/folio/security/advisories/new)

We will acknowledge your report within 48 hours and aim to provide
a fix or mitigation plan within 7 business days.

## What Qualifies

- Unsafe handling of untrusted document content
- Cross-site scripting (XSS) or command injection
- Path traversal or unintended file access
- Exposure of secrets, tokens, or PII
- Broken access control

## What Does Not Qualify

- Denial-of-service attacks
- Social engineering
- Issues in dependencies without a demonstrated exploit path
- Reports from automated scanners without a proof of concept

## Disclosure

We follow coordinated disclosure. Once a fix is released, we will
credit reporters (unless they prefer to remain anonymous) and publish
a security advisory on GitHub.

## Supported Versions

Security fixes are applied to the latest release only. We do not
backport fixes to older versions.
