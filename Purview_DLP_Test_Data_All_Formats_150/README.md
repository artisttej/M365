# Purview DLP Multi-Format Test Data

This package contains 26 Purview DLP test cases with 150 fictional records per test case.

Each test case is supplied in five formats:

- CSV
- DOCX
- PPTX
- XLSX
- TXT

## Data quality

- 3,900 globally unique full names
- 3,900 globally unique physical addresses
- 150 unique sensitive-data values per test case
- No placeholder street-name token remains anywhere in the deliverables
- Country-appropriate address structures are used throughout
- Selected fabricated DLPTest.com card and SSN rows retain row-level provenance
- Applicable identifier formats and checksums were validated

## Folder structure

- `CSV/` - complete eight-column row data
- `DOCX/` - landscape reference tables with identity, address, sensitive data and provenance
- `PPTX/` - cover plus ten readable data slides per test case
- `XLSX/` - formatted, filterable tables containing all eight source columns
- `TXT/` - labelled plain-text records containing all fields

All data is fictional and is intended only for legitimate testing of systems you are authorized to test.
