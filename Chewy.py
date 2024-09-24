import shutil
import os

# Define the source files
chewy_asn = "assets/Chewy 856 ASN - Copy.xlsx"
chewy_label = "assets/Chewy UCC128 Label Request - Copy.xls"

# Define the destination filenames for the copies
chewy_asn_copy = "Finished/Chewy 856 ASN - Copy - Backup.xlsx"
chewy_label_copy = "Finished/Chewy UCC128 Label Request - Copy - Backup.xls"

# Copy the files
shutil.copy(chewy_asn, chewy_asn_copy)
shutil.copy(chewy_label, chewy_label_copy)

print(f"Copied '{chewy_asn}' to '{chewy_asn_copy}'")
print(f"Copied '{chewy_label}' to '{chewy_label_copy}'")