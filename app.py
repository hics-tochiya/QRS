from flask import Flask, render_template, request, redirect, url_for, session, send_file, abort
import openpyxl
import os
import csv
import pdfkit
import zipfile
import secrets
from datetime import datetime, timedelta

# 本物のExcelを裏で操作するためのライブラリ
import win32com.client
import pythoncom  # マルチスレッド（Webサーバー）環境でExcelを安全に動かすために必要

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY') or secrets.token_hex(32)
ADMIN_USERNAME = os.environ.get('QRS_ADMIN_USERNAME', 'admin')
ADMIN_PASSWORD = os.environ.get('QRS_ADMIN_PASSWORD')

# ■設定: wkhtmltopdfのパス (ご自身の環境に合わせてください)
path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
pdf_config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

# 場所リスト
LOCATIONS = [
    '4N1', '4N2', '4N3',
    '4S1', '4S2', '4S3', '4S4',
    '5N1', '5N2', '5N3',
    '5S1', '5S2',
    '6N1', '6N2', '6N3', '6N4',
    '6S1', '6S2', '6S3', '6S4', '6S5',
    '4階フロアー', '5階フロアー', '6階フロアー',
    '旧喫煙室', '図書室', '5階資料室', '6階資料室',
    '職員室', '理事長室','サーバ室'
]

FILE_NAME = '管理台帳＆一覧表.xlsx'
LOG_FILE = 'scan_history.csv'
DOWNLOAD_DIR = 'downloads'

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)


# --- Excel変換機能 ---
def convert_xls_to_xlsx(xls_path, xlsx_path):
    # Webサーバー上でCOM（Windowsの機能）を安全に呼び出すためのおまじない
    pythoncom.CoInitialize()

    # win32comは「絶対パス（C:\...から始まるフルパス）」でないと動かないため変換
    abs_xls = os.path.abspath(xls_path)
    abs_xlsx = os.path.abspath(xlsx_path)

    # 変換先に既にファイルがあれば消しておく
    if os.path.exists(abs_xlsx):
        os.remove(abs_xlsx)

    # パソコンに入っている本物のExcelをバックグラウンドで起動
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # 画面には出さない
    excel.DisplayAlerts = False  # 「上書きしますか？」などの警告ダイアログも無視

    try:
        # 古い .xls を開き、新しい .xlsx 形式 (FileFormat=51) で「名前を付けて保存」
        wb = excel.Workbooks.Open(abs_xls)
        wb.SaveAs(abs_xlsx, FileFormat=51)
        wb.Close()
    finally:
        # 処理が終わったら必ずExcelを終了させる
        excel.Quit()
        pythoncom.CoUninitialize()


# --- データ読み込み機能 ---
def load_inventory_data():
    if not os.path.exists(FILE_NAME):
        return {}

    inventory = {}
    try:
        wb = openpyxl.load_workbook(FILE_NAME, data_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if len(row) <= 2: continue

            mgmt_id = row[2]  # C列 (管理番号)

            if mgmt_id is not None and mgmt_id != "":
                key = str(mgmt_id).strip()
                loc_val = row[7] if len(row) > 7 and row[7] not in (None, "") else "場所未登録"  # H列
                name_val = row[4] if len(row) > 4 and row[4] not in (None, "") else "名称なし"  # E列
                user_val = row[8] if len(row) > 8 and row[8] not in (None, "") else "-"  # I列

                inventory[key] = {
                    'location': str(loc_val).strip(),
                    'name': str(name_val),
                    'user': str(user_val)
                }
    except zipfile.BadZipFile:
        print(f"【エラー】ファイルが壊れています。正しいExcelファイルか確認してください。")
    except Exception as e:
        print(f"【エラー】Excelファイルの読み込みに失敗しました: {str(e)}")

    return inventory


# アプリ起動時にデータをグローバル変数として保持
INVENTORY_DATA = load_inventory_data()


# --- ログ保存機能 ---
def save_scan_log(location, results):
    file_exists = os.path.exists(LOG_FILE)
    with open(LOG_FILE, 'a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(['日時', '作業場所', '判定', '管理番号', '品名', '詳細・使用者', '本来の場所'])
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        for res in results:
            writer.writerow([now_str, location, res['status'], res['code'], res['name'], res['user'], res['reg_loc']])


# --- データ照合ロジック ---
def process_scan_data(current_location, scanned_codes_list):
    results = []
    summary = {'ok': 0, 'ng_place': 0, 'ng_unknown': 0}

    # 1. 読み取ったコードの正誤判定
    for code in scanned_codes_list:
        if code in INVENTORY_DATA:
            item_data = INVENTORY_DATA[code]
            registered_location = item_data['location']

            if current_location == registered_location:
                status = "OK"
                style = "success"
                summary['ok'] += 1
            else:
                status = "場所違い"
                style = "warning"
                summary['ng_place'] += 1

            results.append({
                'code': code, 'name': item_data['name'], 'user': item_data['user'],
                'status': status, 'reg_loc': registered_location, 'style': style
            })
        else:
            summary['ng_unknown'] += 1
            results.append({
                'code': code, 'name': '不明', 'user': '-',
                'status': '未登録', 'reg_loc': '-', 'style': 'danger'
            })

    # 2. 未スキャン（不足）リストの抽出
    missing_items = []
    for key, val in INVENTORY_DATA.items():
        if val['location'] == current_location and key not in scanned_codes_list:
            missing_items.append({'code': key, 'name': val['name'], 'user': val['user']})

    return results, summary, missing_items


# ==========================================
# 画面遷移・ルーティング
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def select_location():
    if request.method == 'POST':
        session['current_location'] = request.form.get('location')
        return redirect(url_for('scan_qr'))

    alert_message = None
    if os.path.exists(FILE_NAME):
        timestamp = os.path.getmtime(FILE_NAME)
        last_modified = datetime.fromtimestamp(timestamp)
        if datetime.now() - last_modified > timedelta(days=60):
            alert_message = f"注意: エクセルファイルが2ヶ月以上更新されていません（最終更新: {last_modified.strftime('%Y/%m/%d')}）"

    grouped_locations = {
        '4階エリア': ['4N1', '4N2', '4N3', '4S1', '4S2', '4S3', '4S4', '4階フロアー'],
        '5階エリア': ['5N1', '5N2', '5N3', '5S1', '5S2', '5階フロアー', '5階資料室'],
        '6階エリア': ['6N1', '6N2', '6N3', '6N4', '6S1', '6S2', '6S3', '6S4', '6S5', '6階フロアー', '6階資料室'],
        'その他': ['旧喫煙室', '図書室', '職員室', '理事長室', 'サーバ室']
    }
    return render_template('select_location.html', grouped_locations=grouped_locations, alert_message=alert_message)


@app.route('/scan', methods=['GET', 'POST'])
def scan_qr():
    if 'current_location' not in session:
        return redirect(url_for('select_location'))

    current_location = session['current_location']

    if request.method == 'POST':
        raw_text = request.form.get('qr_codes', '')
        raw_list = [line.strip() for line in raw_text.splitlines() if line.strip()]
        scanned_codes = list(dict.fromkeys(raw_list))

        session['last_scanned_codes'] = scanned_codes
        results, summary, missing_items = process_scan_data(current_location, scanned_codes)
        save_scan_log(current_location, results)

        return render_template('result.html', current_location=current_location,
                               results=results, summary=summary, missing_items=missing_items)

    return render_template('scan_qr.html', location=current_location)


@app.route('/download_pdf')
def download_pdf():
    scanned_codes = session.get('last_scanned_codes', [])
    current_location = session.get('current_location', '不明な場所')
    if not scanned_codes and not current_location:
        return redirect(url_for('select_location'))

    results, summary, missing_items = process_scan_data(current_location, scanned_codes)
    now_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"result_{current_location}_{now_str}.pdf"
    save_path = os.path.join(DOWNLOAD_DIR, filename)

    rendered_html = render_template('result.html', current_location=current_location,
                                    results=results, summary=summary, missing_items=missing_items, is_pdf=True)
    options = {'page-size': 'A4', 'encoding': "UTF-8", 'no-outline': None}

    try:
        pdfkit.from_string(rendered_html, save_path, configuration=pdf_config, options=options)
        return send_file(save_path, as_attachment=True, download_name=filename)
    except Exception as e:
        return f"<h1>PDF作成エラー</h1><p>{str(e)}</p>"



# --- 保存済みPDFの閲覧・ダウンロード ---
def saved_pdf_path(filename):
    # Windowsの区切り文字も拒否し、保存フォルダ直下のPDFだけを配信する。
    if (not filename or '/' in filename or '\\' in filename
            or not filename.lower().endswith('.pdf')):
        abort(404)
    directory = os.path.realpath(DOWNLOAD_DIR)
    path = os.path.realpath(os.path.join(directory, filename))
    if os.path.dirname(path) != directory or not os.path.isfile(path):
        abort(404)
    return path


@app.route('/pdfs')
def saved_pdfs():
    pdfs = []
    with os.scandir(DOWNLOAD_DIR) as entries:
        for entry in entries:
            if (entry.is_file(follow_symlinks=False)
                    and entry.name.lower().endswith('.pdf')):
                stat = entry.stat()
                pdfs.append({
                    'name': entry.name,
                    'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y/%m/%d %H:%M'),
                    'timestamp': stat.st_mtime,
                    'size_kb': max(1, (stat.st_size + 1023) // 1024),
                })
    pdfs.sort(key=lambda pdf: (pdf['timestamp'], pdf['name']), reverse=True)
    return render_template('saved_pdfs.html', pdfs=pdfs)


@app.route('/pdfs/<filename>/view')
def view_saved_pdf(filename):
    return send_file(saved_pdf_path(filename), mimetype='application/pdf',
                     as_attachment=False, download_name=filename)


@app.route('/pdfs/<filename>/download')
def download_saved_pdf(filename):
    return send_file(saved_pdf_path(filename), mimetype='application/pdf',
                     as_attachment=True, download_name=filename)


# --- 管理者機能 ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if ADMIN_PASSWORD and username == ADMIN_USERNAME and secrets.compare_digest(password or '', ADMIN_PASSWORD):
            session['is_admin'] = True
            return redirect(url_for('admin'))
        else:
            error = 'IDまたはパスワードが間違っています。'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.pop('is_admin', None)
    return redirect(url_for('select_location'))


@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if not session.get('is_admin'):
        return redirect(url_for('login'))

    message = None
    if request.method == 'POST':
        if 'excel_file' not in request.files:
            message = 'ファイルが選択されていません。'
        else:
            file = request.files['excel_file']

            if file.filename == '':
                message = 'ファイルが選択されていません。'

            # --- .xlsx がアップロードされた場合 ---
            elif file.filename.lower().endswith('.xlsx'):
                try:
                    file.save(FILE_NAME)
                    global INVENTORY_DATA
                    INVENTORY_DATA = load_inventory_data()
                    message = '✅ 台帳ファイル(.xlsx)を正常に更新しました。'
                except Exception as e:
                    message = f'❌ 保存エラー: {str(e)}'

            # --- ★ .xls がアップロードされた場合（Excelアプリを使って自動変換） ---
            elif file.filename.lower().endswith('.xls'):
                temp_xls = 'temp_upload.xls'
                try:
                    # 1. 一時的に .xls として保存
                    file.save(temp_xls)

                    # 2. 本物のExcelを操作して .xlsx に変換して保存
                    convert_xls_to_xlsx(temp_xls, FILE_NAME)

                    # 3. 再読み込み
                    INVENTORY_DATA = load_inventory_data()
                    message = '✅ 台帳ファイル(.xls)を【自動で.xlsxに変換して】正常に更新しました。'
                except Exception as e:
                    message = f'❌ 変換エラー: {str(e)}'
                finally:
                    # 一時ファイルを削除して証拠隠滅
                    if os.path.exists(temp_xls):
                        try:
                            os.remove(temp_xls)
                        except:
                            pass
            else:
                message = '❌ エラー：Excelファイル（.xlsx または .xls）のみアップロード可能です。'

    last_modified = "ファイルが存在しません"
    if os.path.exists(FILE_NAME):
        timestamp = os.path.getmtime(FILE_NAME)
        lastanat_modified = datetime.fromtimestamp(timestamp).strftime('%Y年%m月%d日 %H:%M:%S')

    return render_template('admin.html', message=message, last_modified=last_modified)


if __name__ == '__main__':
    # 同じネットワーク内の他端末からアクセスできるように 0.0.0.0 を指定
    app.run(host='0.0.0.0', port=5001, debug=True)