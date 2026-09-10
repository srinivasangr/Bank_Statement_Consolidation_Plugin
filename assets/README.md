`BankStatementTemplate.xlsx` is the workbook a client edits and supplies alongside their statement files.

Do not edit it by hand. It is generated from `src/template.ts`, the same builder the
`bank_statement_create_template` tool uses. To change it:

```bash
npm run build && npm run build:template && npm run verify:template
```

Do not place real bank statements in this repository. Use synthetic samples for tests.
