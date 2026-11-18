# Third-party notices

| Component | Version tested | License | How it is used |
|---|---|---|---|
| Python | 3.12 | PSF-2.0 | runtime |
| PySide6-Essentials / Qt 6 | 6.12.0 | LGPL-3.0 (or GPL) | desktop UI; dynamically linked, unmodified |
| PyMuPDF / MuPDF | 1.27 | **AGPL-3.0** (or commercial) | PDF inspection, merging, page rendering. Compatible with this project's GPL-3.0 license; the full source of this application is published. |
| openpyxl | 3.1.5 | MIT | read-only inspection of OOXML workbooks |
| xlrd | 2.0.2 | BSD-3-Clause | read-only inspection of legacy .xls |
| pywin32 (optional) | 312 | PSF-2.0 | automation of Microsoft Excel when the user has it installed |
| LibreOffice (external, not bundled) | — | MPL-2.0 | conversion engine; the user installs it from libreoffice.org |
| Microsoft Excel (external, optional, not bundled) | — | proprietary | optional conversion engine; never required |

Dev/test only: pytest (MIT), xlwt (BSD) for generating legacy test files.

Before shipping an installer: re-check LGPL obligations for Qt (allow relinking: ship Qt as
separate DLLs), and keep this file in sync with the versions actually bundled.
