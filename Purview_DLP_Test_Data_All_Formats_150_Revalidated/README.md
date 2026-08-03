# Purview DLP SIT Test Data - Revalidated

This revision contains 26 test cases with 150 rows per case in CSV, DOCX, PPTX, XLSX, and TXT formats.

Every row contains:

- a globally unique synthetic full name;
- a globally unique synthetic physical address;
- a pattern/checksum-valid identifier;
- an explicit SIT keyword placed next to the identifier;
- supporting evidence used by Microsoft Purview confidence rules where the built-in definition requires it.

## Credit cards

All 150 card numbers were curated from `artisttej/M365/dummy_credit_card_data.csv`, then independently filtered for a recognized Visa, Mastercard, American Express, Discover, Diners Club, or JCB format and a valid Luhn checksum. Every row also carries the detected brand, a future expiration date, a CVV, and the phrase `Credit Card Number`.

## Built-in coverage

23 cases map to current Microsoft built-in SIT definitions. Three requested cases are not listed in the current built-in catalogue and therefore require a custom SIT:

- US Adoption Taxpayer ID Numbers
- US Preparer Taxpayer ID Numbers
- UK Bank Sort Code and Account Number Combinations

Do not interpret local validation as a guarantee of a particular tenant result. Purview detection also depends on policy scope, configured confidence, exclusions, inspection support for the file type, service updates, and proximity rules. Use `SIT_COVERAGE_MATRIX.csv` and `SIT_VALIDATION_REPORT.json` as the acceptance checklist.

All values are fictional synthetic test data. Do not use them for real transactions, identity claims, tax filings, healthcare, banking, or authentication.
