import os
import re

# Define the path
file_path = r"C:\Users\Lenovo\OneDrive - TEAM\PBŘ_2022_Lorenc\2024_53_Tichý_rekonstrukce_RD_Řečkovice\00b_PBŘ_beta"

# Split the path and get the second-to-last part
path_parts = os.path.normpath(file_path).split(os.sep)
target_part = path_parts[-2] if len(path_parts) >= 2 else None

# Remove the prefix pattern like "2024_53_" using regex
if target_part:
    target_part = re.sub(r"^\d+_\d+_", "", target_part)

# Add the ".docx" extension
target_part_with_extension = f"{target_part}.docx" if target_part else None

# Output the result
print(target_part_with_extension)
