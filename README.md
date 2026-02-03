## Features

### Sort Data by Template List

- Sorts a data table according to the order of names in a separate **template list**.
- Supports:
  - A template column containing the desired name order.
  - A target **name column** in your data that should follow this order.
- Behavior:
  - Creates a backup of both template and target ranges before sorting.
  - Normalizes names (uppercased, spaces removed) and builds a lookup from the template.
  - Uses a temporary helper column to assign sort keys:
    - Exact matches get the template order.
    - Close (approximate) matches:
      - Are assigned the matched template order.
      - Are colored **green** (name contains mistakes but was mapped).
    - Names not in the template:
      - Are pushed to the end with high sort keys.
      - Are colored **red** (not found in template).
  - After sorting, highlights template names:
    - **Blue** if the template name does not appear in the target data (even approximately).
  - Adds an on‑sheet **ColorExplanation** box:
    - Green: Name contains mistakes.
    - Red: Name not in template range.
    - Blue: Missing in target range.

---

### Backup and Undo Changes

- Creates a hidden `__Backup` sheet in each workbook after sorting.
- Backs up a selected range (or two ranges) including:
  - Workbook name, sheet name, range address, and row count.
  - All data, formulas, and formatting.
- Can back up:
  - A **single range**, or
  - **Two ranges** at once (across the same or different workbooks). For the same workbook, the second backup is appended below the first in the same `__Backup` sheet.
- Provides an **Undo Backup** action that:
  - Scans all open workbooks for a `__Backup` sheet.
  - Restores the original data into the backed‑up ranges.
  - Handles two backed‑up blocks if present.
  - Deletes the `__Backup` sheet after a successful restore.
  - Warns if the target sheet is protected or if no backup is found.

---

### Cross‑Window Name Finder


- Uses the value in the active cell as the search term.
- Automatically switches to the next open Excel window.
- Searches the active sheet in that window for the value.
- Search behavior:
  - Performs an exact `Find` first.
  - If not found, scans the used range and:
    - Ignores spaces and case.
    - Allows approximate (“almost”) matches with small spelling differences.
    - Colors the matched cell’s font **blue** when the match is approximate.
- Useful for:
  - Comparing two workbooks side by side.
  - Quickly locating corresponding names/IDs across windows.

---

### Insert into Colored Cells

- Prompts you to select a target range.
- For each cell in that range:
  - If the cell has a fill color (not “no fill”) **and** is empty:
    - Inserts the value `1`.
- Leaves non‑colored cells and already filled cells unchanged.
- Useful for:
  - Turning a colored layout into numeric indicators for formulas, summaries, or pivots.

---


## Usage

### Cross‑Window Name Finder

1. In one Excel window, select a cell containing the name/value you want to find.
2. Run **Switch Window and Find** from the add‑in (e.g., from the Ribbon).
3. The add‑in:
   - Switches to the next Excel window.
   - Searches the active sheet for the same (or similar) value.
4. If found, the matching cell is selected (and may be colored blue for an approximate match); otherwise, you see a “Not Found” message.

---

### Insert into Colored Cells

1. Run **Insert in Colored Cells**.
2. When prompted, select the target range that contains colored cells.
3. The add‑in scans the range and:
   - Fills every **empty, colored** cell with the value `1`.
   - Leaves all other cells unchanged.

---

### Sort Data by Template List

1. Run **Sort By Template**.
2. When prompted:
   1. Select the **template column** that defines the desired order of names.
   2. Select the **name column** in your data that must be sorted according to this template.
3. The add‑in:
   - Backs up the template and the entire data block.
   - Inserts a helper column, computes sort keys, and sorts rows by the template order.
   - Colors names:
     - **Green**: fuzzy‑matched/corrected to a template entry.
     - **Red**: not present in the template.
     - **Blue** (in the template list): missing from the data.
   - Displays a **ColorExplanation** box on the sheet for quick reference.
4. If needed, use **Undo Backup** to return the data and template to their original state.

## Note

This was my first time making a program using VBA. This program does it's job just fine, but can be slow with big ranges, therefore more optimizition needs to be made.
