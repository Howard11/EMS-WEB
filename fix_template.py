import re

# Read the corrupted file
with open(r'c:\Users\icpda\Desktop\工作\EMS\Django\20251203\EMS_Web\ems_site\templates\energy_system_form.html', 'r', encoding='utf-8') as f:
    content = f.read()

# Find where the corruption starts (after line 609) and fix it
# The corruption starts after `newBtn.textContent = "+";`
# We need to remove the corrupt HTML and add proper JavaScript completion

corruption_start = content.find('newBtn.textContent = "+";')
if corruption_start == -1:
    print("Could not find corruption point")
    exit(1)

# Find the end of this line
corruption_start_end = content.find('\n', corruption_start) + 1

# Find where the proper HTML resumes (looking for the closing script tag area)
# We'll rebuild from the proper structure

# Extract the good part before corruption
good_start = content[:corruption_start_end]

# Add the corrected JavaScript completion for addRow and removeRow functions
fixed_js = """            newBtn.classList.remove("btn-danger");
            newBtn.classList.add("btn-success");
            newBtn.setAttribute("onclick", "addRow('" + tableId + "', this)");

            // 插入新列
            tbody.appendChild(newRow);
        }

        function removeRow(tableId, btn) {
            const tbody = document.getElementById(tableId).querySelector("tbody");
            const rows = tbody.querySelectorAll("tr");
            const row = btn.closest("tr");

            // 若只剩一列,不允許刪除
            if (rows.length === 1) {
                alert("至少保留一列資料！");
                return;
            }

            const rowIndex = Array.from(rows).indexOf(row);
            row.remove();

            // 若刪除的是最後一列前一列,需讓它恢復成 "+"
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
    <script>
        // 預填充保存的數據
        const savedDataElement = document.getElementById('saved-data');
        const savedData = savedDataElement ? JSON.parse(savedDataElement.textContent) : {};

        window.addEventListener('DOMContentLoaded', function () {
            if (savedData && Object.keys(savedData).length > 0) {
                // 遍歷所有保存的數據
                for (const [key, value] of Object.entries(savedData)) {
                    if (!value || key === 'csrfmiddlewaretoken') continue;

                    // 處理 checkbox 類型
                    if (key === 'diagram[]') {
                        if (Array.isArray(value)) {
                            value.forEach(v => {
                                const checkbox = document.querySelector(`input[type="checkbox"][value="${v}"]`);
                                if(checkbox) checkbox.checked = true;
                            });
                        } else {
                            const checkbox = document.querySelector(`input[type="checkbox"][value="${value}"]`);
                            if (checkbox) checkbox.checked = true;
                        }
                        // 觸發相應的 toggle 函數
                        if (value.includes && value.includes('無空調系統')) {
                            document.getElementById('noAirSystem').checked = true;
                            toggleAirSystem();
                        }
                        if (value.includes && value.includes('無空壓系統')) {
                            document.getElementById('noAirCompressor').checked = true;
                            toggleAirCompressor();
                        }
                        continue;
                    }

                    // 處理 radio buttons
                    if (key === 'iso14064') {
                        const radio = document.querySelector(`input[type="radio"][name="${key}"][value="${value}"]`);
                        if (radio) radio.checked = true;
                        continue;
                    }

                    // 處理數組類型的輸入 (如照明、空調等表格數據)
                    if (Array.isArray(value)) {
                        const inputs = document.querySelectorAll(`input[name="${key}"]`);
                        value.forEach((v, index) => {
                            if (inputs[index] && v) {
                                inputs[index].value = v;
                            }
                        });
                    } else {
                        // 處理單一值
                        const input = document.querySelector(`input[name="${key}"]`);
                        if (input) {
                            input.value = value;
                        }
                    }
                }
            }
        });
    </script>
</body>

</html>"""

# Write the fixed content
with open(r'c:\Users\icpda\Desktop\工作\EMS\Django\20251203\EMS_Web\ems_site\templates\energy_system_form.html', 'w', encoding='utf-8') as f:
    f.write(good_start + fixed_js)

print("File has been fixed successfully!")
