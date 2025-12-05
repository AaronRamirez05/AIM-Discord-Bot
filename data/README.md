# Excel Data Folder

Place your Excel file (.xlsx or .xls) in this directory.

## Example File Structure

Your Excel file should have structured data with headers in the first row. For example:

| Name | Age | Score | Department |
|------|-----|-------|------------|
| John | 25 | 85 | Sales |
| Jane | 30 | 92 | Marketing |
| Bob | 28 | 78 | IT |

## Configuration

Make sure to update the `EXCEL_FILE_PATH` in your `.env` file to point to your Excel file:

```
EXCEL_FILE_PATH=./data/your_file.xlsx
```

## Supported Formats

- .xlsx (Excel 2007+)
- .xls (Excel 97-2003)
- Multiple sheets are supported
