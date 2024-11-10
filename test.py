import xlwings as xw

# List of sheet information [current sheet name, new sheet name, other value]
d_PU_types = [["CHÚC", "new_name", 2], ["nevýrobní", "new_name_2", 3], ["instalační šachty", "new_name_instalacni", 4], ["instalační šachty", "new_name_instalacni_2", 4]]

# Variable to track if the "instalační šachty" sheet has been duplicated
duplicated_instalacni_shachty = False

# Open the workbook once
with xw.App(visible=False) as app:  # Set visible=True to see what's happening
    workbook = app.books.open("D131_PBŘ_XXX_přílohy.xlsx")

    # Loop through the list in reversed order
    for list_item in reversed(d_PU_types):
        sheet_to_copy_name = list_item[0]  # Current sheet name to copy
        new_sheet_name = list_item[1]  # New sheet name to assign

        # Check conditions for duplicating sheets
        if sheet_to_copy_name == "instalační šachty" and not duplicated_instalacni_shachty:
            duplicated_instalacni_shachty = True  # Mark as duplicated
        elif sheet_to_copy_name == "instalační šachty" and duplicated_instalacni_shachty:
            continue  # Skip further copies if already duplicated

        # Copy the sheet and rename it
        sheet_to_copy = workbook.sheets[sheet_to_copy_name]
        sheet_to_copy.api.Copy(Before=workbook.sheets[0].api)

        # Find and rename the newly copied sheet
        for sheet in workbook.sheets:
            if sheet.name == sheet_to_copy.name + " (2)":  # Default copied sheet name
                sheet.name = new_sheet_name
                break

    workbook.close()
