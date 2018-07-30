# PyInvoice

PyInvoice is a utility for generating invoices from Excel workbooks.

Note that this is a quick-and-dirty solution very closely tailored to my historical timesheet format, so the utility
offers little flexibility for input format. For larger operations, it is probably more sensible to use a proper
invoicing program.

### Features
- Automatic generation of itemized Markdown-formatted invoice
- Minimal command-line parameters - all data is pulled from workbook
- Validation of timesheet sanity
  - Time range overlapping
  - Time range bounds with respect to cycle start/end dates and ticket open/close dates

### Requirements
PyInvoice requires Python 3 and the `openpyxl` package (obtainable via `pip`).

### Usage
PyInvoice requires a specific format in order to properly parse timesheet data. A workbook demonstrating this format
can be found in the **Examples** section.

The format must contain a `META` spreadsheet containing a project name, tracker URL, and cycle start/end dates in cells
`B1` through `B4`, respectively.

Timesheet data will then be parsed from remaining spreadsheets in the workbook. These spreadsheets should contain
time data (a date, start time, and end time) in columns `A` through `C`, starting on row `2`. It must also contain an
hourly rate in decimal format in cell `H2`, a ticket open date in cell `H4`, and a ticket close date in cell `H5` (may
be empty if the ticket is open at time of invoice).

Any information present in the example workbook not described here exists strictly for reference for persons using the
document, and is ignored by the program.

Once a compatible workbook has been created, simply run the script, providing the workbook file as a parameter. An
invoice will be generated in the same directory as the script.

### Example
An example spreadsheet conforming to the PyInvoice format can be found
[here](https://docs.google.com/spreadsheets/d/1DIaW1fchsGCvBjD-FB2-lSgWYDBtUoCsp2D4M_Do9C8/edit). Alternatively, it may
also be found in the `examples` directory of this repository.

### License
PyInvoice is released under the MIT License.
