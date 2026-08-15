---
status: validate-before-real-use
last_verified: 2026-08-15
audience: authorized-operators-and-maintainers
sensitivity: restricted-personal-data
---

# Operator Guide: Otomasi SK ARSC 2026

## Scope

This tool reads active-member data from Excel, fills an official Word template, allocates sequential letter numbers, checks generated numbering, and converts DOCX outputs to PDF using Microsoft Word automation. It does not approve issuance, validate legal wording, or authorize a signatory.

## Sensitive material

The source workbook, requested-name list, generated letters, NIM, department/program, position, signature material, and recipient distribution are restricted. Keep `data/`, `daftar.txt`, `nomor_terakhir.txt`, `docx/`, and `pdf/` out of Git and out of general AI prompts.

## Mandatory first validation

Before real use, copy the project to an isolated test folder or use synthetic input files with the same columns. Do not run `--all` on the real member workbook for a test.

Verify:

1. Required Excel column names match the script, including `Nama Lengkap` and every field used by the template.
2. Every template placeholder is replaced in paragraphs and tables.
3. Department, program, division/bureau, position, date, month Roman numeral, and filename formatting are correct.
4. The first allocated number matches the authorized register.
5. Two sequential synthetic recipients receive unique consecutive numbers.
6. A misspelled name prompts for confirmation and does not silently select a person.
7. Cancel, exception, retry, and partial-generation behavior do not create unnoticed duplicate or skipped numbers.
8. `cek_nomor.py` identifies expected duplicates/gaps.
9. DOCX and converted PDF render correctly in Word/PDF viewers.

## Environment setup

Use the repository batch helpers or create a virtual environment and install `requirements.txt`. PDF conversion requires a compatible Microsoft Word installation and Windows automation; generation of DOCX does not prove PDF conversion will succeed.

## Controlled issuance procedure

1. Obtain authorization for the real member list, template, signatory, issue date, and starting number.
2. Back up `nomor_terakhir.txt`, the current template, and the authoritative external letter register.
3. Put the approved workbook in `data/` under the exact expected filename.
4. Prepare a small reviewed `daftar.txt`; prefer named selection over `--all`.
5. Run generation interactively. Review every fuzzy-name suggestion against a trusted identifier; do not accept on similarity alone.
6. Compare generated numbers with the external register and run the number checker.
7. Open every DOCX sample and all exceptional records; check content and page layout.
8. Convert to PDF and inspect the PDFs, including font substitution, page breaks, and signature area.
9. Obtain issuer approval before distribution.
10. Archive the issuance manifest securely and update the authoritative register; retain personal files only as long as policy requires.

## Recovery

If a run is interrupted, do not edit or regenerate blindly. Preserve the output folder and `nomor_terakhir.txt`, list generated filenames/numbers, compare with the authoritative register, move rejected test outputs to a separate review location, and resume only from an approved next number. Never delete issued evidence just to make numbering contiguous.

## Code-change verification

At minimum, syntax-check scripts and repeat the synthetic two-recipient, fuzzy-match-reject, interrupted-run, duplicate/gap, DOCX-render, and PDF-conversion cases. Any change to numbering or placeholder logic requires operator sign-off.

On 2026-08-15, all four Python scripts parsed successfully. No generation or conversion case was run, so this is syntax evidence only.
