Build a full-stack application called “Data Format Fixer”.

Problem:
Database exports often contain numbers and dates stored as text.
Users should upload an XLSX or CSV file, automatically correct
column data types, and download a corrected Excel file.

Stack:

- Frontend: React + TypeScript.
- Backend: Python FastAPI.
- Excel processing: openpyxl.
- Use Python rules for type detection and conversion.
- The application must work without an AI API key.

Design:

- Simple, clean UI inspired by Nothing’s visual style.
- Off-white background, black text, subtle borders, small red accents.
- No gradients or rainbow colors.
- Responsive layout with drag-and-drop upload.
- Show processing status, detected column types, warnings,
  and a download button.

Workflow:

1. Upload an XLSX or CSV file.
2. Analyze column headers and non-empty values.
3. Detect Integer, Decimal, Date, DateTime, Text, or Mixed.
4. Automatically convert confidently detected values.
5. Preserve uncertain values and display warnings.
6. Automatically download <original_name>\_corrected.xlsx.
7. Keep a manual download button if automatic download is blocked.
8. Let users override column types and regenerate the file.

Conversion rules:

- Convert numeric strings into actual Excel numeric cells.
- Convert date strings into actual Excel date cells.
- Keep character columns as text.
- Preserve blank cells.
- Preserve row and column order.
- Do not delete rows or duplicates.
- Preserve phone numbers, postal codes, account numbers,
  identifiers, leading zeros, and numbers longer than 15 digits
  as text.
- Do not silently round numbers or lose precision.
- Do not guess ambiguous dates such as 01/02/2026.
- Provide an input date-order choice: Auto, Day-first, Month-first.
- Provide output date formats:
  mm/dd/yyyy, dd/mm/yyyy, yyyy-mm-dd.
- Preserve time components in datetime values.
- Do not automatically interpret ordinary numbers as date serials.
- Preserve invalid or uncertain values and report their locations.

XLSX:

- Process every worksheet.
- Preserve sheet names, ordering, formulas, and styles
  wherever practical.
- Do not modify formula cells.
- Never overwrite the original file.

CSV:

- Support common encodings and delimiters.
- Explain that CSV cannot store Excel cell types or formatting.
- Export corrected results as XLSX by default.
- Write formula-like uploaded text as literal text, not formulas.

Optional AI:

- AI is not required for the first version.
- If added, use OpenAI gpt-4o-mini only for ambiguous column
  classification, not for rewriting cell values.
- Keep AI disabled by default.
- Require opt-in before sending minimal data samples to OpenAI.
- Store the API key only in the backend environment.
- Validate AI suggestions with deterministic Python rules.

Backend:

- Validate file contents and supported extensions.
- Add configurable file-size and processing limits.
- Return clear errors.
- Clean up temporary files.
- Isolate uploaded files between users.
- Do not log uploaded data.
- Do not expose API keys to the frontend.

Deliver:

- Working frontend and backend.
- requirements.txt, package.json, .env.example, and README.
- Clear local setup instructions for macOS.
- Tests for numeric conversion, leading zeros, long identifiers,
  ambiguous dates, mixed columns, formulas, and multiple sheets.
- Run relevant tests and the frontend build; report any failures.

I am a beginner.
First inspect the project folder and briefly explain your plan.
Then complete the application, install dependencies, run checks,
and fix errors without waiting for “Done” between steps.

Only pause if an API key is genuinely required. Tell me where
to save it locally and wait for my confirmation. Never ask me
to paste secrets into chat.

When everything is ready, wait for my confirmation before
starting the application. After I confirm, run the frontend
and backend and share the local URL.
