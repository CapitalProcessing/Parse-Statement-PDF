import os
import re

# Path to your Holding folder
folder = r"\\server2\Accounting & Reinsurance\12 Bank (Trust) Statements\2025\Holding"

for filename in os.listdir(folder):
    if filename.lower().endswith(".pdf"):
        name, ext = os.path.splitext(filename)

        # Normalize hyphen spacing or insert if missing
        if "-" in name:
            parts = name.split("-", 1)
            new_name = parts[0].strip() + " - " + parts[1].strip()
        else:
            # Insert hyphen before first number if found
            match = re.search(r"\d", name)
            if match:
                idx = match.start()
                new_name = name[:idx].strip() + " - " + name[idx:].strip()
            else:
                new_name = name  # No change if no number and no hyphen

        new_filename = new_name + ext
        old_path = os.path.join(folder, filename)
        new_path = os.path.join(folder, new_filename)

        if old_path != new_path:
            print(f"Will rename: {filename} → {new_filename}")
            # Uncomment next line to actually rename
            os.rename(old_path, new_path)
