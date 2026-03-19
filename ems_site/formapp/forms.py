from django import forms

class EMSForm(forms.Form):
    # --- 基本公司資料 ---
    company_name = forms.CharField(label="公司名稱", max_length=100)
    tax_id = forms.CharField(label="統一編號", max_length=20, required=False)
    company_address = forms.CharField(label="公司地址", max_length=200, required=False)
    industry_type = forms.CharField(label="產業類別", max_length=50, required=False)
    building_type = forms.CharField(label="建物分類", max_length=50, required=False)

    capital = forms.DecimalField(label="資本額 (千元)", required=False)
    revenue = forms.DecimalField(label="營業額 (千元)", required=False)
    employees = forms.IntegerField(label="員工人數", required=False)

    contact_name = forms.CharField(label="聯絡人", max_length=50)
    contact_title = forms.CharField(label="職稱", max_length=50, required=False)
    phone = forms.CharField(label="電話", max_length=50, required=False)
    email = forms.EmailField(label="E-mail", required=False)

    # --- 電力使用 ---
    total_power_usage = forms.DecimalField(label="全區電力使用總量 (度/年)", required=False)
    total_power_cost = forms.DecimalField(label="全區電力使用費用 (萬元/年)", required=False)

    # --- 非電力能源使用量 ---
    gas_usage = forms.DecimalField(label="天然氣 (立方公尺)", required=False)
    lpg_usage = forms.DecimalField(label="液化石油氣 (公噸)", required=False)
    gasoline_usage = forms.DecimalField(label="汽油 (公秉)", required=False)
    diesel_usage = forms.DecimalField(label="柴油 (公秉)", required=False)
    other_usage_1 = forms.CharField(label="其他能源 (1)", required=False)
    other_usage_2 = forms.CharField(label="其他能源 (2)", required=False)

    # --- 非電力能源費用 ---
    gas_cost = forms.DecimalField(label="天然氣 (萬元/年)", required=False)
    lpg_cost = forms.DecimalField(label="液化石油氣 (萬元/年)", required=False)
    gasoline_cost = forms.DecimalField(label="汽油 (萬元/年)", required=False)
    diesel_cost = forms.DecimalField(label="柴油 (萬元/年)", required=False)
    other_cost_1 = forms.DecimalField(label="其他能源 (1) (萬元/年)", required=False)
    other_cost_2 = forms.DecimalField(label="其他能源 (2) (萬元/年)", required=False)

    # --- 各類能源熱值 ---
    gas_heat = forms.DecimalField(label="天然氣", required=False)
    lpg_heat = forms.DecimalField(label="液化石油氣", required=False)
    gasoline_heat = forms.DecimalField(label="汽油", required=False)
    diesel_heat = forms.DecimalField(label="柴油", required=False)
    other_heat_1 = forms.DecimalField(label="其他 (1)", required=False)
    other_heat_2 = forms.DecimalField(label="其他 (2)", required=False)
