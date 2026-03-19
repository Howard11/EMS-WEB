from django.shortcuts import render
from django.http import JsonResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
from django.utils.encoding import escape_uri_path
from .forms import EMSForm
from .word_generator import generate_energy_system_word
import json, os

@csrf_exempt
def ems_form(request):
    data_path = os.path.join(os.path.dirname(__file__), 'data')
    os.makedirs(data_path, exist_ok=True)
    json_path = os.path.join(data_path, 'ems_basic_data.json')

    if request.method == 'POST':
        form = EMSForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data

            # 把所有日期或非字串的內容轉成字串（避免 datetime / Decimal 錯誤）
            for k, v in data.items():
                if hasattr(v, "strftime"):
                    data[k] = v.strftime("%Y-%m-%d")
                elif not isinstance(v, (str, int, float, bool, type(None))):
                    data[k] = str(v)

            # ✅ 覆蓋舊資料（不再 append）
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4, default=str)

            return JsonResponse({'status': 'ok', 'msg': '✅ 資料已保存成功！'})
        else:
            return JsonResponse({'status': 'error', 'msg': '❌ 表單驗證失敗'})
    else:
        # ✅ GET 時預先讀取 JSON，讓網頁一開就有舊資料
        if os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    saved_data = json.load(f)
            except Exception:
                saved_data = {}
        else:
            saved_data = {}

        # 把舊資料傳入表單初始值
        form = EMSForm(initial=saved_data)

    return render(request, 'ems_form.html', {'form': form})


@csrf_exempt
def building_form(request):
    """第二頁：建物資料表單"""
    data_path = os.path.join(os.path.dirname(__file__), 'data')
    save_path = os.path.join(data_path, 'building_data.json')
    os.makedirs(data_path, exist_ok=True)

    if request.method == "POST":
        # 獲取表單數據，處理多選框 (lists)
        data = {}
        for key, value in request.POST.lists():
            if key != 'csrfmiddlewaretoken' and key != 'action':
                if len(value) == 1:
                    data[key] = value[0]
                else:
                    data[key] = value

        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4, default=str)

        return JsonResponse({"status": "ok", "building msg": "✅ 建物資料已保存成功！"})

    # ✅ GET 時讀取舊資料並傳遞給模板
    saved_data = {}
    if os.path.exists(save_path):
        try:
            with open(save_path, 'r', encoding='utf-8') as f:
                saved_data = json.load(f)
        except Exception:
            saved_data = {}
    
    # 轉換為JSON字符串供JavaScript使用
    saved_data_json = json.dumps(saved_data)
    
    return render(request, 'building_form.html', {
        'saved_data': saved_data, 
        'saved_data_json': saved_data_json
    })

@csrf_exempt
def energy_system_form(request):
    """第三頁：建築能效系統與設備資料"""
    from django.http import FileResponse
    from .word_generator import generate_energy_system_word
    import datetime
    
    data_path = os.path.join(os.path.dirname(__file__), 'data')
    save_path = os.path.join(data_path, 'energy_system_data.json')
    os.makedirs(data_path, exist_ok=True)

    if request.method == "POST":
        # 獲取表單數據
        data = {}
        for key, value in request.POST.lists():
            if key != 'csrfmiddlewaretoken' and key != 'action':
                if len(value) == 1:
                    data[key] = value[0]
                else:
                    data[key] = value
        
        # 保存JSON文件
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4, default=str)

        # 根據動作決定回傳內容
        action = request.POST.get('action', 'export')
        
        if action == 'save':
            return JsonResponse({"status": "ok", "msg": "✅ 資料已保存成功！"})
            
        # ✅ 超強合併所有頁面的數據
        all_data = {}
        
        # 定義可能的路徑 (順序決定優先權，後者覆蓋前者)
        possible_paths = [
            os.path.join(data_path, 'building_data.json'),
            os.path.join(data_path, 'energy_system_data.json'),
            os.path.join(data_path, 'ems_basic_data.json'),
        ]
        
        for p in possible_paths:
            if os.path.exists(p):
                try:
                    with open(p, 'r', encoding='utf-8') as f:
                        file_data = json.load(f)
                        if isinstance(file_data, dict):
                            all_data.update(file_data)
                except Exception: continue
        
        # 最後更新當前 POST 數據 (權重最高)
        all_data.update(data)
        
        # 清理舊的 Word 檔案 (確保每次匯出不會在伺服器上堆積無用檔案)
        import glob
        old_files = glob.glob(os.path.join(data_path, '輔導單位基礎資料調查表_*.docx'))
        for old_file in old_files:
            try:
                os.remove(old_file)
            except Exception as e:
                print(f"Warning: Could not remove old file {old_file}: {e}")
        
        # 生成Word文檔 (使用模板)
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        word_filename = f'輔導單位基礎資料調查表_{timestamp}.docx'
        word_path = os.path.join(data_path, word_filename)
        
        try:
            generate_energy_system_word(all_data, word_path)
            
            # 返回Word文件下載
            response = FileResponse(open(word_path, 'rb'), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            response['Content-Disposition'] = "attachment; filename*=utf-8''{}".format(escape_uri_path(word_filename))
            return response
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            return JsonResponse({"status": "error", "msg": f"生成Word文檔失敗: {str(e)}"})

    # ✅ GET 時讀取舊資料並傳遞給模板
    saved_data = {}
    if os.path.exists(save_path):
        try:
            with open(save_path, 'r', encoding='utf-8') as f:
                saved_data = json.load(f)
        except Exception:
            saved_data = {}
    
    # 轉換為JSON字符串供JavaScript使用
    saved_data_json = json.dumps(saved_data)
    
    return render(request, 'energy_system_form.html', {'saved_data': saved_data, 'saved_data_json': saved_data_json})