# -*- coding: utf-8 -*-
import os

path = r'c:\Users\moshei1\OneDrive - Verifone\Desktop\TIP\STFPNOW\בדיקות\Dashboard_App.py'

with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# indices are 0-based
# We want to keep 0 to 612 (lines 1 to 613)
# We want to skip 613 to 843 (lines 614 to 844)
# We want to keep 844 onwards (lines 845 to end)

kept_lines = lines[:613] + lines[844:]

with open(path, 'w', encoding='utf-8') as f:
    f.writelines(kept_lines)

print("Surgical deletion of lines 614-844 complete.")
