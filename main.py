from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.worksheet.datavalidation import DataValidation


def apply_column_validation(
    file_path: str,
    main_sheet_name: str,
    lists_sheet_name: str,
    header_row: int = 1,
    max_validation_row: int = 1000,
) -> None:
    """
    Applies data validation (dropdown list) to columns in the main sheet
    by automatically matching columns with the same header name in the lists sheet.
    """
    wb = load_workbook(file_path)

    if main_sheet_name not in wb.sheetnames:
        raise ValueError(f"Main sheet '{main_sheet_name}' not found.")
    if lists_sheet_name not in wb.sheetnames:
        raise ValueError(f"Lists sheet '{lists_sheet_name}' not found.")

    ws_main = wb[main_sheet_name]
    ws_lists = wb[lists_sheet_name]

    # 1. Map columns from the lists sheet: { "column_name": (col_index, last_filled_row) }
    lists_map = {}
    for col_idx in range(1, ws_lists.max_column + 1):
        header = ws_lists.cell(row=header_row, column=col_idx).value
        if header is not None:
            normalized_name = str(header).strip()

            # Find the last populated row in this specific column
            last_row = header_row
            for row_idx in range(ws_lists.max_row, header_row, -1):
                if ws_lists.cell(row=row_idx, column=col_idx).value is not None:
                    last_row = row_idx
                    break

            # Only include if there is at least 1 item below the header
            if last_row > header_row:
                lists_map[normalized_name] = (col_idx, last_row)

    # Clear existing validations in the main sheet to avoid duplicates on re-runs
    ws_main.data_validations.dataValidation.clear()

    # Format sheet name with single quotes if it contains spaces (e.g., 'Lookup Lists')
    quoted_lists_sheet = quote_sheetname(lists_sheet_name)

    # 2. Iterate through the main sheet columns and create validation when headers match
    for col_idx in range(1, ws_main.max_column + 1):
        main_header = ws_main.cell(row=header_row, column=col_idx).value
        if main_header is None:
            continue

        column_name = str(main_header).strip()

        if column_name in lists_map:
            list_col_idx, last_list_row = lists_map[column_name]
            list_col_letter = get_column_letter(list_col_idx)
            main_col_letter = get_column_letter(col_idx)

            # Formula pointing to the exact range of that column in the lists sheet
            # Example: ='Lists'!$A$2:$A$15
            range_formula = (
                f"={quoted_lists_sheet}!${list_col_letter}${header_row + 1}:"
                f"${list_col_letter}${last_list_row}"
            )

            dv = DataValidation(
                type="list",
                formula1=range_formula,
                allow_blank=True,
            )
            dv.errorTitle = "Invalid Value"
            dv.error = f"Please select a valid value for '{column_name}'."
            dv.promptTitle = column_name
            dv.prompt = "Select an item from the list."

            ws_main.add_data_validation(dv)

            # Target range in the main sheet where validation will be applied (e.g., A2:A1000)
            target_range = f"{main_col_letter}{header_row + 1}:{main_col_letter}{max_validation_row}"
            dv.add(target_range)

            print(
                f"[OK] Column '{column_name}' ({target_range}) -> "
                f"List at {range_formula}"
            )

    wb.save(file_path)
    print(f"\nFile successfully saved at: {file_path}")

if __name__ == "__main__":
    file_name = "sample_validated.xlsx"

    # 2. Apply data validation by matching headers across both sheets
    apply_column_validation(
        file_path=file_name,
        main_sheet_name="Records",
        lists_sheet_name="Lists",
        header_row=1,
        max_validation_row=500,
    )
