import os

def find_path(name, path):
    for root, dirs, files in os.walk(path):
        if name in dirs or name in files:
            return root
    return None

root_dir = r'C:\Users\Oliver'
target_name = 'Test345564.pdf'
path = find_path(target_name, root_dir)
if path:
    print(f"Found at: {path}")
else:
    print(f"{target_name} not found")


