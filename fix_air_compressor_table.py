import re

# Read the backup file  
with open(r'c:\Users\icpda\Desktop\工作\EMS\Django\20251203\EMS_Web\ems_site\templates\energy_system_form.html.backup', 'r', encoding='utf-8') as f:
    content = f.read()

# Find the air compressor table section and fix its header
# Looking for the pattern with rowspan="2" and colspan="5"
old_header = '''                <table class="table table-bordered align-middle text-center" id="airCompressorTable">
                    <thead class="table-light">
                        <tr>
                            <th rowspan="2">項目</th>
                            <th colspan="5">使用率</th>
                        </tr>
                        <tr>
                            <th>1</th>
                            <th>2</th>
                            <th>3</th>
                            <th>4</th>
                            <th>5</th>
                        </tr>
                    </thead>'''

# New header structure that matches the image: 項目 | 使用率 | 1(%) | 2(%) | 3(%) | 4(%) | 5(%)
new_header = '''                <table class="table table-bordered align-middle text-center" id="airCompressorTable">
                    <thead class="table-light">
                        <tr>
                            <th>項目</th>
                            <th>使用率</th>
                            <th>1<br>%</th>
                            <th>2<br>%</th>
                            <th>3<br>%</th>
                            <th>4<br>%</th>
                            <th>5<br>%</th>
                        </tr>
                    </thead>'''

# Replace the header
content = content.replace(old_header, new_header)

# Write the fixed file
with open(r'c:\Users\icpda\Desktop\工作\EMS\Django\20251203\EMS_Web\ems_site\templates\energy_system_form.html', 'w', encoding='utf-8') as f:
    f.write(content)

print("Air compressor table header fixed!")
print("Before: Item | [Usage Rate spanning 1-5]")
print("After: Item | Usage Rate | 1(%) | 2(%) | 3(%) | 4(%) | 5(%)")
