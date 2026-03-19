import re

# Read the corrupted file
with open(r'c:\Users\icpda\Desktop\工作\EMS\Django\20251203\EMS_Web\ems_site\templates\energy_system_form.html', 'r', encoding='utf-8') as f:
    content = f.read()

# Pattern to find the corrupted section starting after line 610
corruption_pattern = r'(newBtn\.textContent = "\+";)\s*\r?\n\s*<tr>.*?</div >\s*\r?\n\s*<script>'

# Replacement text with proper JavaScript completion
replacement = r'''\1
            newBtn.classList.remove("btn-danger");
            newBtn.classList.add("btn-success");
            newBtn.setAttribute("onclick", "addRow('" + tableId + "', this)");

            // 插入新列
            tbody.appendChild(newRow);
        }

        function removeRow(tableId, btn) {
            const tbody = document.getElementById(tableId).querySelector("tbody");
            const rows = tbody.querySelectorAll("tr");
            const row = btn.closest("tr");

            // 若只剩一列，不允許刪除
            if (rows.length === 1) {
                alert("至少保留一列資料！");
                return;
            }

            const rowIndex = Array.from(rows).indexOf(row);
            row.remove();

            // 若刪除的是最後一列前一列，需讓它恢復成 "+"
            const remainingRows = tbody.querySelectorAll("tr");
            if (rowIndex === remainingRows.length) {
                const lastRow = remainingRows[remainingRows.length - 1];
                const lastBtn = lastRow.querySelector("button");
                lastBtn.textContent = "+";
                lastBtn.classList.remove("btn-danger");
                lastBtn.classList.add("btn-success");
                lastBtn.setAttribute("onclick", "addRow('" + tableId + "', this)");
            }
        }
    </script>

    {{ saved_data|json_script:"saved-data" }}
    <script>'''

# Apply the fix using regex
fixed_content = re.sub(corruption_pattern, replacement, content, flags=re.DOTALL)

# Also remove the duplicate functions that appear later
# Find and remove duplicate script section starting from line ~629
duplicate_pattern = r'\r?\n\s*<script>\s*\r?\n\s*function toggleAirSystem\(\) {[\s\S]*?(?=\s*</script>\s*\r?\n\s*</body>)'
fixed_content = re.sub(duplicate_pattern, '', fixed_content, count=1)

# Write the fixed content
with open(r'c:\Users\icpda\Desktop\工作\EMS\Django\20251203\EMS_Web\ems_site\templates\energy_system_form.html', 'w', encoding='utf-8') as f:
    f.write(fixed_content)

print("File has been fixed successfully!")
