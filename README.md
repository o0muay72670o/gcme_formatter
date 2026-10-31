# Engineering Excel Form Converter

A Python project for VS Code. Each Excel file can have a different layout. Python
extracts cell content and merged ranges, AI proposes a structured mapping, and
Python builds the new form using your Master Header.

## Start in VS Code

The local `.venv` is already configured on this computer. Open the project folder,
select `.venv\Scripts\python.exe` using **Python: Select Interpreter**, then open
**Terminal → New Terminal**.

```powershell
.\.venv\Scripts\python.exe main.py inspect "input/ECM-05-008 AT 4 Ceiling Closure Check Sheet.xlsx"
```

This command is local and makes **no AI request**. It writes the extracted content
to `output/reports/`. The supplied sample and Master have been copied into `input/`
and `templates/`; your Downloads originals are unchanged.
These workbooks are local only and excluded from Git. After cloning the repository,
copy your source `.xlsx` files into `input/` and your Master into `templates/`.

You can also press F5 and select **Inspect sample (no AI cost)**. VS Code will
recommend Microsoft's Python and debugger extensions if they are missing.

## Enable AI mapping

1. Copy `.env.example` to `.env`.
2. Fill in `OPENAI_API_KEY` and `OPENAI_MODEL` locally. Choose a model available to
   your API account that supports Responses API Structured Outputs. No model is
   silently selected. Do not put the key in source code or send it through chat.
3. Run a plan first:

```powershell
python main.py plan "input/ECM-05-008 AT 4 Ceiling Closure Check Sheet.xlsx" --sheets "ECM-05-008 AT4_P1" "ECM-05-008 AT4_P2"
```

`--sheets` is optional. It is useful for this sample because its workbook contains
multiple duplicate representations and a blank active sheet. Without it, all
nonempty worksheets are sent for mapping. Inspect first to see their names.

The plan command sends the extracted text and layout metadata plus the Master
Header context to `https://api.openai.com/v1/responses`. It does not upload the
whole workbook or send images. This is a billable API call. There is no separate
AI call for each formatting operation. Keys come from environment variables or
your local `.env` file.

Review the saved `*.plan.json` and `*.report.json`, then generate:

```powershell
python main.py convert "input/ECM-05-008 AT 4 Ceiling Closure Check Sheet.xlsx" --sheets "ECM-05-008 AT4_P1" "ECM-05-008 AT4_P2"
```

The identical request reuses its cached plan. To use an edited plan, add
`--plan "output/reports/ECM-05-008 AT 4 Ceiling Closure Check Sheet.plan.json"`.
Keep the source, settings and `--sheets` selection the same. Edited plans are
validated before generation.

## Run a folder

```powershell
python main.py batch input
# Optional: process two independent files concurrently.
python main.py batch input --workers 2
```

Put `.xlsx` source files in `input/`. The Master lives separately in `templates/`.
Outputs and reports go to `output/`. Existing output files are skipped; use
`--output output/run2` for a fresh run. Files that fail do not stop the rest of
the batch. A batch with failures returns a nonzero exit code.

Cached AI plans use a hash of the source bytes, Master bytes, selected sheets,
configuration, prompt version and model. The cache is for the **same input**;
it does not guess that two different files have the same layout. Delete a cache
entry deliberately if you want another billable analysis of that same input.

## What follows your prompt

- Renames the active source worksheet to `Old` (or a collision-safe suffix).
- Adds a new sheet based on the filename. Excel's 31-character limit requires
  truncation, and duplicate names receive a suffix.
- Copies the Master Header's cells, merged ranges, formatting and row heights.
- Applies the 0.44 cm grid and A4 portrait with the specified margins: left 1.6,
  right 0.6, top/bottom 1.9, header/footer 0.8 cm.
- Builds variable body tables from the AI plan, including fields absent from the
  Header. Tables can have different columns for checks, data and signatures.
- Sets a print area and Fit to 1 page wide by 1 page tall.

`config.json` holds the settings. Header rows are 1–10 and the grid ends at AP for
your supplied Master. Change these if a different Master is used. Physical Excel
column widths depend on the workbook's Normal font and display rounding; 0.44 cm
is a target, not a guarantee for every workbook font.

## Checks and current limits

This is a working first version, **not a validated converter for every engineering
form**. The local extraction, validation, caching and workbook generation are
tested. Live AI mapping requires your API configuration and has not been run here.

Every nonempty extracted source cell must be represented or explicitly excluded
with a reason. Invalid IDs and attempts to overwrite Master labels are rejected.
This detects unaccounted cells, but it does not prove that AI preserved their
meaning. Review exclusions, wording, units, blank response boxes and sign-off
fields before using a generated engineering form.

One-page settings do not prove readability or actual printed page count. The
report includes an estimated scale and flags a small estimated font. Native Excel
print verification is not automated in this version; check Print Preview before
accepting a form. A large source form may need human layout decisions.

Image-only text, embedded drawings, charts, handwritten marks and complex Excel
controls are not interpreted. Add a vision/OCR stage later if the actual file set
needs it. Images/charts detected in the source trigger review notes. `openpyxl`
can lose unsupported shapes or controls in an output copy. Keep original inputs;
the program never overwrites them. `.xls`, `.xlsm`, password-protected files and
dependent formulas requiring an automatic sheet-rename repair are not supported.
Master Header formulas are rejected rather than copied with uncertain references.
Master branding uses copied cells; floating logos are not copied automatically.
Original formulas remain on source sheets; body formulas are preserved as literal
text, not recalculated or executed.

There are no automatic API retries after timeouts, to avoid silently repeating
possibly billable requests. Reports record usage returned by the API; cached
requests record a cache hit. Concurrency reduces waiting time, not token usage.

## Tests

```powershell
python -m unittest discover -s tests -v
```

On another computer install Python 3.11 or newer, then:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

This computer's environment uses the existing bundled Python with system packages
because no standalone Python was installed. It can be recreated with a standard
Python installation later. No global packages or editor extensions were installed.

API implementation reference:
[OpenAI Structured Outputs](https://developers.openai.com/api/docs/guides/structured-outputs).
