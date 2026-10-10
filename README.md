# Equipment Management System (EMS)

A single-file, offline desktop app for **equipment maintenance scheduling** and **spare parts inventory**, built with Python, Tkinter and SQLite. Originally designed for SMT / electronics manufacturing lines, but generic enough for any workshop, lab or plant.

> One `.py` file. No server. No account. No internet required.

<!-- Add screenshots to a /docs folder and uncomment:
![Dashboard](docs/dashboard.png)
-->

## Features

- **Dashboard** - KPI cards (active equipment, overdue tasks, due soon, parts below safety stock, inventory value), plus live maintenance and low-stock alerts
- **Maintenance tracker** - recurring tasks with frequency, auto-calculated next-due dates, urgency colouring, search and filters
- **Calendar view** - monthly calendar of upcoming and overdue tasks
- **Task completion** - log who/when/how long/findings; record **spare parts consumed** and have stock deducted automatically
- **Bulk complete** - select multiple tasks (Ctrl/Shift+click) and complete them at once
- **Equipment registry** - codes, models, serial numbers, status, notes
- **Spare parts inventory** - safety stock, reorder quantity, cost, supplier, lead time, compatible equipment, check-in / check-out with full transaction history
- **Reorder list** - everything at or below safety stock, grouped by supplier, with estimated cost and Excel export
- **Attachments** - attach images and documents to equipment and parts, with inline image preview
- **Documents library** - general store for manuals, SOPs and drawings, with tags
- **Export Centre** - styled Excel exports for history, task status, inventory, transactions, reorder list, or everything in one workbook
- **Excel-driven configuration** - bulk-load equipment, tasks and parts from a spreadsheet
- **Right-click menus** and **double-click to edit** across the main tables
- **SQLite (WAL mode)** - robust local storage

## Requirements

- Python **3.9+**
- Tkinter (bundled with Python on Windows and macOS; on Debian/Ubuntu: `sudo apt install python3-tk`)
- [`openpyxl`](https://pypi.org/project/openpyxl/) (required)
- [`Pillow`](https://pypi.org/project/pillow/) (optional - only for image previews)

```bash
pip install -r requirements.txt
```

## Quick start

```bash
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
pip install -r requirements.txt
python equipment_manager.py
```

On first launch the app creates, next to the script:

| File / folder      | Purpose                                              |
|--------------------|------------------------------------------------------|
| `ems_data.db`      | SQLite database (all your data)                      |
| `ems_config.xlsx`  | Auto-generated template config with example data     |
| `attachments/`     | Uploaded files                                       |
| `exports/`         | Excel exports                                        |

### Try it with the sample data

A ready-made dummy dataset is in [`sample_data/ems_config.xlsx`](sample_data/ems_config.xlsx). To load it:

1. Run the app once, then close it.
2. Copy `sample_data/ems_config.xlsx` over the generated `ems_config.xlsx`.
3. Delete `ems_data.db` (and any `ems_data.db-wal` / `ems_data.db-shm`) for a clean start.
4. Launch the app again - the config is loaded automatically when the database is empty.

Later changes to the spreadsheet can be applied from **Settings -> Reload Config from Excel**.

> The sample's "last done" dates are fixed when the file was generated, so some tasks will show as overdue - that is intentional and shows off the alert colours.

## Configuration file (`ems_config.xlsx`)

Four sheets. Row 1 is a header row; data starts on row 2. **Column order matters** (headers are not read by name).

**Settings** - `key`, `value`

| key               | meaning                                  | example                    |
|-------------------|------------------------------------------|----------------------------|
| `facility_name`   | Title shown on the dashboard             | `Demo Factory - SMT Line 1`|
| `alert_soon_days` | "Due soon" window in days                | `7`                        |
| `default_user`    | Pre-filled "Done by" / "Performed by"    | `Technician`               |
| `currency`        | Currency label for costs                 | `USD`                      |

**Equipment** - `code`, `name`, `category`, `location`, `model`, `serial_no`, `notes`

**MaintenanceTasks** - `equipment_code`, `task_name`, `frequency_days`, `last_done` (YYYY-MM-DD), `assigned_to`, `procedure_notes`

**SpareParts** - `part_number`, `description`, `category`, `location`, `unit`, `qty_on_hand`, `safety_stock`, `reorder_qty`, `unit_cost`, `supplier`, `lead_time_days`, `compatible_equip` (comma-separated equipment codes), `notes`

**How reloading behaves**

- Rows are matched on **equipment `code`**, **part number**, or **equipment + task name**, and updated if they already exist.
- Reloading does **not** overwrite `qty_on_hand` for parts that already exist, so day-to-day stock movements are preserved.
- Tasks, equipment or parts removed from the spreadsheet are **not** deleted from the database.

## Where is my data?

Everything lives in the folder containing the script. **Back up `ems_data.db` and `attachments/` regularly.** To move to another machine, copy those along with the script.

## Packaging into an executable (optional)

The repo intentionally ships only the source file. If you want a standalone Windows executable you can build one yourself with [PyInstaller](https://pyinstaller.org/):

```bash
pip install pyinstaller
pyinstaller --noconsole --onedir --name EMS equipment_manager.py
```

Notes:

- If a file named `equipment.ico` sits next to the script it is used as the window icon. For the executable's icon add `--icon equipment.ico`.
- The app stores its database and attachments relative to the script location. When packaged, make sure that location is writable and persistent (prefer `--onedir`; with `--onefile` the app unpacks to a temporary folder on every launch).

## Limitations

- Designed for **local, single-user** use. There is no authentication or role-based permission system.
- SQLite WAL mode tolerates several local readers, but placing the database on a network share with multiple simultaneous writers is **not recommended**.
- No automatic e-mail / push reminders; alerts appear in the dashboard, status bar and window title.
- Tested primarily on Windows 10/11. Linux and macOS should work but have had less testing.

## Contributing

Issues and pull requests are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md). Please **do not attach real company data** (databases, config files, photos of equipment) to issues.

## License

Released under the [MIT License](LICENSE).
