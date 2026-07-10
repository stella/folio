# Security Audit

Run a security-focused code audit with a generic checklist first, then
layer on repo-specific risks.

## Instructions

1. **Start with repo context**:
   - what kind of system is this?
   - what data does it handle?
   - what are the trust boundaries?
   - what would be high-impact failures here?

2. **Check for hardcoded secrets**:
   - API keys
   - tokens
   - passwords
   - private connection strings

3. **Review authentication and authorization**:
   - are protected routes actually protected?
   - are object/resource ownership checks enforced server-side?
   - are role checks done at the boundary, not just in the UI?

4. **Review input handling**:
   - validation present for user-controlled inputs
   - no obvious SQL injection, command injection, XSS, or path traversal
   - file paths, URLs, and object keys are sanitized appropriately

5. **Review file and storage access** when relevant:
   - upload validation
   - download/presign access checks
   - safe retention/deletion behavior

6. **Review network and session behavior**:
   - sane session expiration and invalidation
   - rate limits where brute force is plausible
   - external requests have timeouts and failure handling
   - CORS is not broader than necessary

7. **Review dependency risk**:

   ```bash
   bun audit
   ```

   Use the repo's equivalent audit command if it does not use Bun. Also check
   GitHub security or Dependabot alerts when available.

8. **Review AI-specific risks** when applicable:
   - prompt injection exposure
   - unsafe interpolation of user content into system instructions
   - cross-tenant or cross-document leakage through retrieval/context assembly

9. **Layer in domain-specific risks**:
   - finance, healthcare, legal, infra, auth, multi-tenant SaaS, and internal tools
     all have different sharp edges
   - explicitly call out the domain assumptions you used

10. **Report findings by severity**:

- Critical
- High
- Medium
- Low

For each finding include:

- file and line
- issue
- likely impact
- recommended fix

11. **If there are no findings**, say so explicitly and mention what was checked
    plus any residual gaps in verification.
