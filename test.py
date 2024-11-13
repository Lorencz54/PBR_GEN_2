import win32com.client as win32

# Open Word application
word = win32.Dispatch('Word.Application')
word.Visible = False  # Set to True if you want to see the Word application

# Open the source and destination documents
source_doc = word.Documents.Open(r'C:\Users\Lenovo\OneDrive\Projekty\Skamba\sablony\PBR\D131_PBŘ_nevýrobní_TZ.docx')
destination_doc = word.Documents.Open(r'C:\Users\Lenovo\OneDrive\Projekty\Skamba\sablony\PBR\table_test.docx')

# Get the table from the source document
source_table = source_doc.Tables(2)  # Assuming the table you want to copy is the first one

# Copy the table
source_table.Range.Copy()

# Move to the end of the destination document and paste the table
destination_doc.Content.InsertAfter('\n')  # Adds a newline to move the cursor
destination_doc.Content.Paste()  # Paste the table

# Save the destination document
destination_doc.SaveAs(r'C:\Users\Lenovo\OneDrive\Projekty\Skamba\sablony\PBR\table_test.docx')

# Close the documents
source_doc.Close()
destination_doc.Close()

# Quit Word application
word.Quit()


