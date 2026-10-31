# Engineering Excel Form Converter

Use **Codex in this chat to understand each form**, then Python to create the Excel
output. Different files can have different layouts. **No API key is needed.**

Python does not connect itself to your Codex conversation. It prepares local JSON
requests; Codex reads them in this workspace and writes plans; Python applies the
plans. This is assisted batch processing through a chat you start or continue.

## Use it in VS Code

The Python environment is already configured on this computer. Select
`.venv\Scripts\python.exe` with **Python: Select Interpreter**.

1. Put source `.xlsx` files in `input/`.
2. Put the Master in `templates/Master File.xlsx`.
3. Open **Terminal → New Terminal** and run:

```powershell
.\.venv\Scripts\python.exe main.py prepare input
```

4. In **this Codex chat with the project workspace**, say:

> Process the prepared requests in jobs using the Master Header. Create the plans
> and apply them. Use this chat for the AI analysis; do not use an API.

Codex can do the preparation too. You can simply say:

> Process the Excel files in input using this project's no-API workflow.

5. Find the new workbooks in `output/`, with per-file reports in `output/reports/`.
   Review engineering wording, exclusions, sign-off fields and Excel Print Preview.

To run the application step yourself after Codex has created the plans:

```powershell
.\.venv\Scripts\python.exe main.py apply jobs
```

F5 also offers **1. Prepare folder for Codex chat (no API)** and
**2. Apply Codex chat plans (no API)**. Preparation alone does not generate plans;
ask Codex in between those two steps. There is no `.env` setup in this workflow.

## Sample file

The sample and Master are already copied into the local input/template folders.
They are excluded from Git; after cloning, copy your own workbooks there.

The sample includes original pages, helper tables, duplicate exports and a blank
active sheet. Select the two complete table representations for analysis:

```powershell
.\.venv\Scripts\python.exe main.py prepare "input/ECM-05-008 AT 4 Ceiling Closure Check Sheet.xlsx" --sheets "Table 1" "Table 2"
```

For an inventory without creating a chat request:

```powershell
.\.venv\Scripts\python.exe main.py inspect "input/ECM-05-008 AT 4 Ceiling Closure Check Sheet.xlsx"
```

Without `--sheets`, prepare extracts all nonempty source sheets. Codex must account
for duplicates explicitly; the program does not silently drop their content.

## How the handoff works

- `prepare FILE_OR_FOLDER` writes `jobs/<name>-<hash>/request.json`: source cell
  IDs, text, locations, merges, Master labels, allowed header targets and plan schema.
- Codex writes a `plan.json` envelope beside each request. It contains the exact
  source_hash and a completed plan object. AGENTS.md guides future project chats.
- `apply jobs` validates each plan, rechecks the source/Master/settings fingerprint,
  and builds the output. You can also apply one request.json path.
- Files without plans report waiting_for_chat. Invalid plans fail without stopping
  other requests. Existing outputs are skipped. Use `--output output/run2` for a
  fresh output folder.

Re-running prepare does not overwrite an existing plan. Unchanged jobs can reuse
their saved plans without another AI analysis. A changed workbook, Master, selected
sheets or configuration gets a new fingerprint. Old requests for changed files are
rejected. Keep only requests you want to process in jobs, or apply specific paths.

`batch input` and `plan FILE` also default to preparing chat requests.
`convert FILE --plan ENVELOPE.json` accepts a saved envelope. None of these makes an
API call by default. Optional API mode requires explicitly adding `--api` plus an
API key/model; it is not required for the chat workflow.

## Excel behavior

config.json contains your prompt's settings:

- Master Header rows 1–10 through AP, with the form title in blank cell P4.
- 0.44 cm target grid width, A4 portrait, left margin 1.6 cm, right 0.6 cm,
  top/bottom 1.9 cm, header/footer 0.8 cm.
- Active source sheet renamed to Old (or a collision-safe suffix).
- New sheet named from the filename, truncated to Excel's 31-character limit.
- Variable body tables for unmatched fields, inspection items, remarks and signatures.
- Print area and Fit to one page wide by one page tall.

Master labels are preserved. Actual header values come from referenced source
cells. plan.title supplies the normalized form title. Original inputs are never
overwritten. The output copy retains the other source sheets.

## Review and current limits

Every extracted nonempty cell must be represented or excluded with a reason.
Unknown IDs, missing coverage, malformed tables and attempts to overwrite Master
labels are rejected. These checks do not prove semantic correctness: review the
engineering wording, units, exclusions, response fields and sign-off sections.

One-page settings do not prove actual printed page count or legibility. Reports
estimate shrinkage; native Excel print verification is not automated. Review Print
Preview. Large forms may need human layout decisions. Physical column widths also
depend on the workbook's Normal font and Excel rounding.

Image-only text, handwritten marks, charts, drawings and complex controls are not
interpreted by the extractor. Images/charts trigger review notes. Floating Master
logos are not copied. openpyxl may lose unsupported shapes or controls in output
copies; keep originals. .xls, .xlsm, protected files and Master Header formulas are
not supported. Active sheets with dependent formulas/named ranges are blocked from
automatic renaming. Body formulas are literal text, not executed or recalculated.

This workflow uses the usage limits of your Codex chat. It does not consume API
credits or independently run unattended AI sessions. Processing many unrelated
layouts still takes AI analysis; Python handles repeated Excel formatting.

## Tests and installation

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

On another computer, install Python 3.11+ and create the environment:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

This computer's local environment uses the available bundled Python runtime.
The environment, workbooks, jobs, reports and secrets are excluded from Git.
No paid API calls are needed to install, test or use the chat workflow.
