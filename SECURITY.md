# Security Policy

## Security considerations

werkmap writes a spreadsheet file and never reads one. It parses nothing, so it has no parser to attack, and it turns no string into code: the suite runs under `--disallow-code-generation-from-strings`, a Chromium page runs it under `default-src 'none'; script-src 'self'`, and the source is scanned on every run.

Do not treat werkmap as a sanitiser for the data you put in it. It escapes what it writes into XML, and it strips the characters XML cannot carry, so author text reaches a cell as text rather than as markup. That is the whole of its claim. It does not decide which data belongs in a document.

Anyone who can reach the produced file can read every value in it. Cells hold what you passed, and the document properties hold what you passed as `meta`. Keep out of a workbook what its readers may not see.

A spreadsheet application is the other half of the threat model, and it is not this package:

- **Text beginning with `=` is written as text**, never as a formula, so there is no formula-injection surface here. An application that re-interprets a pasted value is outside what this package can prevent.
- **Image bytes are stored verbatim.** werkmap checks that you named `png` or `jpeg` and never decodes the file. A malformed image is a question for whatever opens the workbook.
- **A document is only as trustworthy as its source.** Nothing in the format stops a reader acting on what it finds.

Memory is the one resource worth bounding. The whole document is assembled before `bytes()` returns, so a row count you do not control is a row count you should cap.

## Reporting a vulnerability

Do not open a public GitHub issue for a security vulnerability.

Use [GitHub's private vulnerability form](https://github.com/getquario/werkmap/security/advisories/new).

Include the affected code, its impact, and steps that reproduce the issue. Tell us whether and how to credit you.

We do not accept AI slop reports.

Keep the report private while we investigate and prepare a fix.
