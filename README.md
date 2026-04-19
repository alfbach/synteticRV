# Synthetic RVTools (.xlsx)

Small **Flask** web UI plus a **Python** generator that produces a synthetic **RVTools-style** Excel workbook from a template file named **`customer.xlsx`** (same sheet layout as a typical RVTools export: `vInfo`, `vCPU`, `vDisk`, and so on).

## What it does

- **S / M / L**: take the first *N* VMs from `vInfo` (about **120 / 400 / 1000**), then keep only rows on other sheets that still belong to those VMs (plus consistent host, cluster, datastore and multipath filtering where applicable).
- **XL**: keep **all** VMs from the template.
- **XXL**: same as XL, then **add synthetic VMs** so the VM count grows by **40%** compared to the template (`ceil(original_vm_count × 0.4)` extra VMs). Each extra VM is built by **cloning** an existing template VM’s rows on every sheet that has a `VM` column (paths and free-text fields replace the seed VM name with a unique internal token, then the usual anonymisation maps everything to `SYN-VM-xxxxx`, etc.).

Infrastructure sheets without a `VM` column are **not** duplicated for clones (hosts/clusters stay as in the template subset), which matches the idea of “more guest VMs on the same estate”.

## Requirements

- Python **3.10+** recommended  
- `customer.xlsx` placed next to `app.py`

## Quick start

```bash
./start.sh
```

Then open **http://127.0.0.1:8765** in your browser.

Manual equivalent:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

## UI behaviour

- **Output filename**: sanitised and forced to end with `.xlsx` if needed.
- **Choose folder** (Chromium / Edge): uses the **File System Access API** to write the generated file into a directory you pick.
- Without a folder: the app returns the file as a **normal download**.
- **Server path** (optional): if you run Flask locally and send a JSON body with `export_dir`, the server writes `{export_dir}/{filename}`. Only use this on machines you trust; the path is not validated beyond creating directories.

## API

`POST /api/generate` with JSON:

```json
{
  "size": "m",
  "filename": "out.xlsx",
  "export_dir": "/optional/absolute/path"
}
```

If `export_dir` is omitted or empty, the response is the **file bytes** (`Content-Disposition: attachment`). If `export_dir` is set, the response is JSON including `written_path` and metadata (`vm_count`, and for XXL also `original_vm_count`, `synthetic_clones`).

## Performance

Large templates (many thousands of rows per sheet) can take **tens of seconds** to process, especially **XXL**, because rows are duplicated and the workbook is written with **openpyxl**.

## License

No license file is bundled; add one if you distribute this project.
