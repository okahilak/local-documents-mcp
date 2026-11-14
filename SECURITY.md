# Security Policy

## Reporting a vulnerability

Please report security vulnerabilities privately using GitHub's [private vulnerability reporting](https://github.com/okahilak/local-documents-mcp/security/advisories/new). Do not open a public issue for security problems.

Useful details to include:

- What the issue allows (for example, reading or changing files outside the allowed folders)
- Steps or a sample file to reproduce it
- Your operating system and the Local Documents version

Reports are acknowledged as soon as possible and investigated with care. Once a fix is released, the issue is disclosed in a security advisory, with credit to the reporter if they want it.

## Scope

Examples of issues that are in scope:

- Accessing, creating, moving or modifying files outside the allowed folders (including via `..`, symbolic links or Windows junctions)
- Overwriting or deleting files when the tool promises not to
- Crashes or code execution triggered by a crafted PDF, Word or Excel file

## Supported versions

Security fixes are made to the latest release only.
