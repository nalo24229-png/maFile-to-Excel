import os
import json
import re
import shutil
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Настройка путей строго по русским названиям папок
DATA_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "Вписывать данные сюда"))

TXT_ACCOUNTS = os.path.join(DATA_DIR, "accounts.txt")
PROXIES_FILE = os.path.join(DATA_DIR, "proxies.txt")
MAFILES_DIR = os.path.join(DATA_DIR, "mafiles")

OUTPUT_ROOT = os.path.abspath(os.path.join(BASE_DIR, "..", "Ваши готовые аккаунты"))
DIR_TOGETHER = os.path.join(OUTPUT_ROOT, "Все аккаунты вместе")
DIR_SEPARATE = os.path.join(OUTPUT_ROOT, "Каждый по отдельности")

OUTPUT_EXCEL_TOGETHER = os.path.join(DIR_TOGETHER, "market_bots.xlsx")

# ИСПРАВЛЕНО: Скрипт ищет файл "Шаблон.xlsx" внутри папки "Под капотом"
TEMPLATE_FILE = os.path.join(BASE_DIR, "Шаблон.xlsx")

WINRAR_PATH = r"C:\Program Files\WinRAR\WinRAR.exe"

def parse_accounts(file_path):
    acc_data = {}
    if not os.path.exists(file_path): 
        print(f"[-] ОШИБКА: Файл '{file_path}' не найден!")
        return acc_data
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.replace("\n", "").replace("\r", "")
            if not line or ":" not in line: continue
            parts = line.split(":")
            if len(parts) >= 2:
                login = parts[0].strip()
                password = parts[1].strip()
                acc_data[login] = password
    return acc_data

def parse_proxies(file_path):
    if not os.path.exists(file_path): 
        print(f"[-] ПРЕДУПРЕЖДЕНИЕ: Файл прокси '{file_path}' не найден. Сборка идет без прокси.")
        return []
    res = []
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.replace("\n", "").replace("\r", "")
            if line: res.append(line.strip())
    return res

def pack_xml_to_xlsx(xml_content, target_xlsx_path):
    base_target_dir = os.path.dirname(target_xlsx_path)
    temp_dir = os.path.join(base_target_dir, "temp_build_xlsx_dir")
    temp_worksheets = os.path.join(temp_dir, "xl", "worksheets")
    os.makedirs(temp_worksheets, exist_ok=True)
    
    with open(os.path.join(temp_worksheets, "sheet1.xml"), "w", encoding="utf-8") as f:
        f.write(xml_content)
        
    shutil.copyfile(TEMPLATE_FILE, target_xlsx_path)
    
    current_dir = os.getcwd()
    os.chdir(temp_dir)
    
    winrar_cmd = [
        WINRAR_PATH,
        "a",
        os.path.abspath(target_xlsx_path),
        os.path.join("xl", "worksheets", "sheet1.xml")
    ]
    subprocess.run(winrar_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    os.chdir(current_dir)
    shutil.rmtree(temp_dir, ignore_errors=True)

def main():
    if not os.path.exists(WINRAR_PATH) or not os.path.exists(TEMPLATE_FILE):
        print("[-] ОШИБКА: Системные файлы конфигурации или WinRAR не найдены!")
        return

    txt_passwords = parse_accounts(TXT_ACCOUNTS)
    proxies_list = parse_proxies(PROXIES_FILE)

    if not os.path.exists(MAFILES_DIR):
        print("[-] ОШИБКА: Папка с мафайлами mafiles отсутствует!")
        return

    files = [f for f in os.listdir(MAFILES_DIR) if f.endswith(".maFile")]
    files.sort()
    if not files:
        print("[-] ОШИБКА: В папке mafiles не найдено ни одного файла .maFile!")
        return

    xml_header = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://openxmlformats.org" '
        'xmlns:r="http://openxmlformats.org" '
        'xmlns:mc="http://openxmlformats.org" mc:Ignorable="x14ac" '
        'xmlns:x14ac="http://microsoft.com">\n'
        '  <dimension ref="A1:E1000"/>\n'
        '  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>\n'
        '  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>\n'
        '  <sheetData>\n'
        '    <row r="1" rCustomHeight="1">\n'
        '      <c r="A1" t="inlineStr"><is><t>login</t></is></c>\n'
        '      <c r="B1" t="inlineStr"><is><t>password</t></is></c>\n'
        '      <c r="C1" t="inlineStr"><is><t>alias</t></is></c>\n'
        '      <c r="D1" t="inlineStr"><is><t>proxy</t></is></c>\n'
        '      <c r="E1" t="inlineStr"><is><t>maFile</t></is></c>\n'
        '    </row>\n'
    )
    
    xml_footer = '  </sheetData></worksheet>'
    xml_together_body = ""

    os.makedirs(DIR_TOGETHER, exist_ok=True)
    os.makedirs(DIR_SEPARATE, exist_ok=True)
    
    valid_idx = 0
    for file_name in files:
        try:
            file_path = os.path.join(MAFILES_DIR, file_name)
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read().strip()

            json_match = re.search(r"\{.*\}", content, re.DOTALL)
            if not json_match: continue
            mafile_json = json.loads(json_match.group(0))

            login = mafile_json.get("account_name", "")
            if not login: continue

            session = mafile_json.get("Session", {}) or {}
            steam_id = session.get("SteamID", "") or mafile_json.get("steamid", "")
            
            if steam_id:
                account_id = int(steam_id) - 76561197960265728
                calculated_trade_url = f"https://steamcommunity.com{account_id}"
            else:
                calculated_trade_url = ""

            password = txt_passwords.get(login, "") or mafile_json.get("password", "")
            safe_content = content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', '&quot;').replace("'", '&apos;')

            if proxies_list:
                proxy = proxies_list[valid_idx % len(proxies_list)]
                proxy_val = proxy.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', '&quot;').replace("'", '&apos;')
            else:
                proxy_val = calculated_trade_url

            row_num_together = valid_idx + 2
            row_xml_together = (
                f'    <row r="{row_num_together}" rCustomHeight="1" xml:space="preserve">\n'
                f'      <c r="A{row_num_together}" t="inlineStr"><is><t>{login}</t></is></c>\n'
                f'      <c r="B{row_num_together}" t="inlineStr"><is><t>{password}</t></is></c>\n'
                f'      <c r="C{row_num_together}" t="inlineStr"><is><t>{login}</t></is></c>\n'
                f'      <c r="D{row_num_together}" t="inlineStr"><is><t>{proxy_val}</t></is></c>\n'
                f'      <c r="E{row_num_together}" t="inlineStr" xml:space="preserve"><is><t>{safe_content}</t></is></c>\n'
                f'    </row>\n'
            )
            xml_together_body += row_xml_together

            bot_folder = os.path.join(DIR_SEPARATE, login)
            os.makedirs(bot_folder, exist_ok=True)

            shutil.copyfile(file_path, os.path.join(bot_folder, file_name))

            info_txt_path = os.path.join(bot_folder, "info.txt")
            with open(info_txt_path, "w", encoding="utf-8") as info_f:
                info_f.write(f"Логин аккаунта: {login}\n")
                info_f.write(f"Пароль аккаунта: {password}\n")
                info_f.write(f"Авто-ссылка обмена: {calculated_trade_url}\n")
                info_f.write(f"Привязанный прокси: {proxy if proxies_list else 'Отсутствует'}\n")

            row_xml_single = (
                f'    <row r="2" rCustomHeight="1" xml:space="preserve">\n'
                f'      <c r="A2" t="inlineStr"><is><t>{login}</t></is></c>\n'
                f'      <c r="B2" t="inlineStr"><is><t>{password}</t></is></c>\n'
                f'      <c r="C2" t="inlineStr"><is><t>{login}</t></is></c>\n'
                f'      <c r="D2" t="inlineStr"><is><t>{proxy_val}</t></is></c>\n'
                f'      <c r="E2" t="inlineStr" xml:space="preserve"><is><t>{safe_content}</t></is></c>\n'
                f'    </row>\n'
            )
            single_xml_full = xml_header + row_xml_single + xml_footer
            single_xlsx_path = os.path.join(bot_folder, f"market_bot_{login}.xlsx")
            
            pack_xml_to_xlsx(single_xml_full, single_xlsx_path)
            valid_idx += 1
        except Exception as e:
            print(f"[-] Ошибка обработки аккаунта: {e}")

    together_xml_full = xml_header + xml_together_body + xml_footer
    pack_xml_to_xlsx(together_xml_full, OUTPUT_EXCEL_TOGETHER)
    
    print("\n[+] УСПЕХ! Процесс генерации полностью завершен.")
    print(f"[ инфо ] Всего ботов успешно упаковано в таблицы: {valid_idx}")

if __name__ == "__main__":
    main()