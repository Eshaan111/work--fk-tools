# Flipkart PDF print middleware

For installation on another computer, see `MACHINE_SETUP.md` and run
`Run Machine Setup.cmd` after the physical TTP-244 Pro has been installed.

## Saved printer connection

Physical printer settings recorded on 31 July 2026:

```text
Queue:        TSC TTP-244 Pro
Driver:       TSC TTP-244 Pro
Port:         USB001
Port monitor: Dynamic Print Monitor
USB device:   USB\VID_1203&PID_0172\000001
USB printer:  USBPRINT\TSCTTP-244_PRO\6&2B02C22&2&USB001
```

Flipkart capture routing:

```text
Queue shown to Flipkart/QZ: PDFCreator
Driver:                      TSC TTP-244 Pro
Port:                        FlipkartRawCapturePort
Destination:                 RAW TCP 127.0.0.1:9100
Original PDFCreator queue:   PDFCreator Original -> pdfcmon
```

After reconnecting the printer, Windows should bind it to `USB001`
automatically. Verify that `TSC TTP-244 Pro` shows status `Normal`, then start
the middleware:

```powershell
cd "C:\Users\ESHAAN\HAKUR\work--fk-tools\order-grouper"
python raw_print_capture.py
```

If Windows does not recreate the physical queue but the driver and `USB001`
port exist, recreate it from an elevated PowerShell window:

```powershell
Add-Printer -Name "TSC TTP-244 Pro" -DriverName "TSC TTP-244 Pro" -PortName "USB001"
```

The same settings are stored in `printer-connection.json` for machine-readable
reference.

Flipkart currently labels its jobs as `QZ Tray Raw Print`. For those jobs, use
the raw capture path:

```text
Flipkart -> QZ Tray -> Flipkart Raw Capture printer
         -> TCP 127.0.0.1:9100 -> raw_print_capture.py
         -> print-jobs/raw
```

The raw capture worker detects PDF, ZPL, TSPL, EPL, ESC/POS, PostScript, and
unknown binary jobs. A real PDF is copied into `print-jobs/incoming` for the OCR
worker. Printer-language jobs remain in `print-jobs/raw` so their exact format
can be handled without corrupting the bytes.

Process a captured ZPL job into individual labels:

```powershell
python zpl_job_processor.py <captured.zpl> print-jobs/labels/<job-name>
```

Send a captured batch unchanged to the physical printer:

```powershell
python print_zpl_batch.py <captured.zpl> --printer "TSC TTP-244 Pro"
```

## Automatic batch-print dialog

Double-click `Launch Flipkart Batch Printer.cmd`. The dialog owns TCP port
9100, waits for a Flipkart/QZ batch, extracts each label, sorts complete label
blocks, and submits the sorted ZPL batch once to `TSC TTP-244 Pro`. Sorting is
color (ICE, BEIGE, BLACK, WHITE), then fit (BAGGY, PLAIN), then size ascending;
unclassified labels are tagged MIX and sent last. A separate large-text divider
label is inserted before every category (for example, `ICE-28` or
`WHITE-PLAIN-32`). No captured Flipkart label content, layout, or format is
modified. A SKU containing two or more recognized size tokens is treated as
MIX; a trailing `_39` means size 32 only when the SKU is not already ambiguous.
Each divider also shows the number of order labels in that category and the
seller alias (`PRABHU` for PRABHU ENTERPRISES and `SEEMA` for SEEMA
ENTERPRISES). If both appear in one category, both aliases are shown.

Do not run `raw_print_capture.py` at the same time as the dialog because only
one process can listen on port 9100.

This splits every printable `^XA...^XZ` block and creates one ZPL source, JSON
record, and CSV row per label. Order ID, SKU, product, quantity, payment type,
service type, carrier, packaging instruction, seller, AWB, fit, color, size and
sort tag are extracted from the original ZPL text without OCR. The resulting
print sequence is saved as `sort-order.json`. Add preview rendering only when
visual inspection is explicitly needed.

Start raw capture before printing:

```powershell
python raw_print_capture.py
```

Then select `Flipkart Raw Capture` in Flipkart/QZ Tray.

## Rendered PDF path

PDFCreator is only suitable when QZ submits a rendered Windows print job. It
cannot convert a QZ raw job into PDF.

The rendered path is:

```text
Flipkart -> QZ Tray -> PDFCreator
         -> PDFCreator COM queue -> print-jobs/incoming
         -> print_middleware.py -> OCR JSON
```

## Install and configure the physical printer

1. Install the official TSC Windows driver and confirm that a Windows test page
   prints correctly on the physical `TSC TTP-244 Pro`.
2. In the TSC printer preferences, create/select the exact label stock in use.
   A common Flipkart shipping label is 100 x 150 mm, but use the physical stock
   actually loaded in the printer.
3. Install PDFCreator. Its standard `PDFCreator` virtual printer is sufficient;
   the worker consumes its COM job queue directly.
4. Select `PDFCreator` as the printer in Flipkart/QZ Tray.

## 2. Install the worker dependencies

Install Python packages:

```powershell
python -m pip install -r requirements.txt
```

Install Tesseract OCR for Windows and SumatraPDF. The example configuration
expects:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
C:\Program Files\SumatraPDF\SumatraPDF.exe
```

## 3. Configure and run

Adjust `print-middleware.json` if the printer name, executable paths, OCR
language, or print settings differ.

Keep `"printing": { "enabled": false }` for the first test. Run:

```powershell
python print_middleware.py
```

Print one label to `PDFCreator`. The worker should create:

```text
print-jobs/archive/<job>.pdf
print-jobs/archive/<job>.ocr.json
```

Inspect the PDF and OCR JSON. Once they are correct, change `printing.enabled`
to `true`. SumatraPDF then silently submits the archived job to the named TSC
queue. Its `paper=auto,noscale` settings preserve the PDF page dimensions; the
TSC driver stock should match those dimensions.

## Operational behavior

- A PDF must stop changing for `stable_seconds` before it is claimed.
- The worker must already be running when a job is sent to the `PDFCreator`
  queue. Only one program should consume that COM queue.
- OCR uses a 300 DPI render, which is appropriate for source labels even though
  the TTP-244 Pro prints at 203 DPI.
- Successful jobs and OCR data go to `archive`.
- Failed jobs and a matching `.error.txt` file go to `failed`.
- Printing is only attempted after PDF extraction and OCR succeed.
