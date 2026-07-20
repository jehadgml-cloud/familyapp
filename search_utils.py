import re

filepath = "app.js"
with open(filepath, "r", encoding="utf-8") as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "function openEditMemberModal" in line or "function deleteMember" in line:
        print(f"Line {i+1}: {line.strip()}")
