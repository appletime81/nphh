import os

for root, dirs, files in os.walk("."):
    for file in files:
        if file.endswith(".py"):
            file_path = os.path.join(root, file)
            os.system(f"black {file_path}")
