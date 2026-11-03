"""
PM 週會議程產生腳本（2026-08-31 全面重寫）

用法：
    python generate_agenda.py <config.json>

config.json 由 Claude 在互動確認完 SKILL.md 的 Step1~3 之後產生，
本腳本只做機械式的檔案操作，不做語意判斷（比對哪個PM對應哪個重點專案、
排除名單是否異動等仍由 Claude 在動腳本前先跟使用者確認完畢）。

config.json 結構：
{
  "source_xlsx": "來源檔案絕對路徑",
  "output_xlsx": "輸出檔案絕對路徑",
  "tue_date": "2026-09-01",           // 本週週二日期
  "thu_date": "2026-09-03",           // 本週週四日期
  "listed_projects": [                // Claude比對「重點專案」找出的總表原始列文字(含開頭編號)，
                                       // 用來清空重寫總表D欄「列管」標記
    "(2)機關戊_資料整合服務案：成員K、成員I、成員J、成員N、成員L"
  ],
  "upper_order": ["PM_A", "PM_B", "PM_C", "PM_D", "PM_E"],  // 週四上半場PM順序
  "tue_order_override": null,         // 若要調整週二PM順序，給完整名字陣列；否則null=沿用總表順序
  "lower_order_override": null,       // 同上，週四下半場
  "absentees": [                      // 本次會議缺席者，會被移到當天清單最後並加註記
    {"name": "PM_Q", "day": "thu", "type": "business_trip"},
    // type: business_trip | leave
    // 週二出差且要實際移去週四下半場報告時，加 "moved_to": "thu"：
    // 週二清單仍會列出他/她並加註記，週四下半場清單會額外把他/她的週二項目整組搬過去
    {"name": "PM_L", "day": "tue", "type": "business_trip", "moved_to": "thu"}
  ],
  "duty_rotation": null                // 若本次要重排未來值日生才給，否則null
  // duty_rotation範例：
  // {
  //   "weekday": "四",
  //   "start_date": "2026-09-17",
  //   "end_date": "2026-12-31",
  //   "cycle": ["PM_E", "PM_O", "PM_H", "PM_P", "PM_C"]
  // }
}
"""
import sys
import json
import re
import datetime as dt
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

THIN = Side(style='thin')
BORDER_ALL = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
FILL_NAME = PatternFill(fill_type='solid', fgColor='FFFF9900')
NO_FILL = PatternFill(fill_type=None)

FONT_TITLE = Font(name='微軟正黑體', size=12, bold=True)
FONT_HEADER = Font(name='微軟正黑體', size=13, bold=True)
FONT_NAME = Font(name='微軟正黑體', size=11, bold=False)
FONT_PROJ = Font(name='微軟正黑體', size=11, bold=False)
FONT_LISTED = Font(name='微軟正黑體', size=12, bold=True, color='FFFF0000')

MIN_PER_ITEM = 10  # 上半場每個列管專案的報告分鐘數

ANNOTATIONS = {
    ('tue', 'business_trip'): '(本週出差，改週四報告)',
    ('thu', 'business_trip'): '(本週出差，改上傳簡報摘要)',
    ('tue', 'leave'): '(本週請假)',
    ('thu', 'leave'): '(本週請假)',
}


def strip_num(raw):
    return re.sub(r'^\(\d+\)', '', raw)


def parse_section(ws, start_row, end_row):
    """把總表(正常報告版)的一個區段(週二或週四)解析成 [(name, [item,...]), ...]"""
    result, cur_name, cur_items = [], None, []
    for r in range(start_row, end_row + 1):
        a = ws.cell(row=r, column=1).value
        if a is None:
            continue
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        d = ws.cell(row=r, column=4).value
        if b is None and c is None and not str(a).startswith('('):
            if cur_name:
                result.append((cur_name, cur_items))
            cur_name, cur_items = a, []
        else:
            cur_items.append({'raw': a, 'type': b, 'goal': c, 'listed': (d == '列管')})
    if cur_name:
        result.append((cur_name, cur_items))
    return result


def find_section_bounds(ws):
    """自動找總表裡『週二報告(正常)』『週四報告(正常)』兩個區段的起訖列"""
    tue_start = thu_start = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if a == '週二報告(正常)':
            tue_start = r + 2  # 跳過標題列+欄頭列
        elif a == '週四報告(正常)':
            thu_start = r + 2
    tue_end = thu_start - 3 if thu_start else ws.max_row
    thu_end = ws.max_row
    return tue_start, tue_end, thu_start, thu_end


def apply_absentees(order, groups_dict, day, absentees):
    """把缺席者移到清單最後，並回傳 {name: 註記文字或None}"""
    notes = {}
    movers = []
    remain = []
    for name in order:
        hit = next((a for a in absentees if a['name'] == name and a['day'] == day), None)
        if hit:
            notes[name] = ANNOTATIONS[(day, hit['type'])]
            movers.append(name)
        else:
            notes[name] = None
            remain.append(name)
    return remain + movers, notes


def find_rota_row(rota, target_date):
    for r in range(1, rota.max_row + 1):
        a = rota.cell(row=r, column=1).value
        if isinstance(a, dt.datetime) and a.date() == target_date:
            return r
    return None


def clear_fill_border_font(cell):
    cell.value = None
    cell.font = Font()
    cell.fill = NO_FILL
    cell.border = Border()


def style_row(ws, r, is_name=False):
    for col in (1, 2, 3):
        c = ws.cell(row=r, column=col)
        c.border = BORDER_ALL
        if is_name:
            c.fill = FILL_NAME


def write_line(ws, r, text, bold=True):
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(name='Arial', size=11, bold=bold)
    return r + 1


def main(config_path):
    with open(config_path, encoding='utf-8') as f:
        cfg = json.load(f)

    wb = openpyxl.load_workbook(cfg['source_xlsx'], data_only=False)
    zb = wb['專案總表(正常報告版)']

    rota = wb['「115年(會議輪值表)」'] if '「115年(會議輪值表)」' in wb.sheetnames else None
    if rota is None and not all(k in cfg for k in ('chair', 'duty_tue', 'duty_thu', 'boss_chairs_thu')):
        raise RuntimeError(
            "來源檔沒有「115年(會議輪值表)」分頁，無法自動查主席/值日生。"
            "請在config補上手動欄位：chair, duty_tue, duty_thu, boss_chairs_thu(布林值)"
        )

    if '專案議程' in wb.sheetnames:
        agenda = wb['專案議程']
    else:
        agenda = wb.create_sheet('專案議程')
        agenda.column_dimensions['A'].width = 76.5
        agenda.column_dimensions['B'].width = 12.16
        agenda.column_dimensions['C'].width = 11.16

    tue_date = dt.date.fromisoformat(cfg['tue_date'])
    thu_date = dt.date.fromisoformat(cfg['thu_date'])

    # ---------- Stage A: 總表D欄「列管」標記重寫 ----------
    for r in range(1, zb.max_row + 1):
        if zb.cell(row=r, column=4).value == '列管':
            zb.cell(row=r, column=4).value = None
            zb.cell(row=r, column=4).font = Font()
    listed_set = set(cfg['listed_projects'])
    marked_rows = []
    for r in range(1, zb.max_row + 1):
        a = zb.cell(row=r, column=1).value
        if a in listed_set:
            c = zb.cell(row=r, column=4, value='列管')
            c.font = Font(name='微軟正黑體', size=12, bold=True, color='FFFF0000')
            marked_rows.append(r)
    missing = listed_set - {zb.cell(row=r, column=1).value for r in marked_rows}
    if missing:
        raise RuntimeError(f"總表找不到以下列管項目原文，請檢查文字是否一致：{missing}")
    print(f"[StageA] 總表D欄已重寫，{len(marked_rows)}筆列管標記")

    # ---------- 解析總表週二/週四區段 ----------
    tue_start, tue_end, thu_start, thu_end = find_section_bounds(zb)
    tue = parse_section(zb, tue_start, tue_end)
    thu = parse_section(zb, thu_start, thu_end)
    # 已離職/不再報告的PM（config `exclude_pms`），總表還沒刪時先在這裡排除
    # 專案改掛別人（config `reassign`，2026-11-03新增，例：離職者專案分給他人）
    # 格式 [{"from": "PM_H", "to": "PM_K", "keyword": "機關庚"}]；項目接在to名下最後、重新編號
    for rs in cfg.get('reassign') or []:
        src_list = tue if any(n == rs['from'] for n, _ in tue) else thu
        src = dict(src_list)[rs['from']]
        hits = [it for it in src if rs['keyword'] in it['raw']]
        if len(hits) != 1:
            raise RuntimeError(f"reassign找不到唯一項目：{rs}")
        it = hits[0]
        src.remove(it)
        dst = dict(tue).get(rs['to']) if any(n == rs['to'] for n, _ in tue) else dict(thu).get(rs['to'])
        if dst is None:
            raise RuntimeError(f"reassign找不到接手PM：{rs['to']}")
        it['raw'] = f"({len(dst) + 1}){strip_num(it['raw'])}"
        dst.append(it)
    excluded = set(cfg.get('exclude_pms') or [])
    tue = [(n, its) for n, its in tue if n not in excluded]
    thu = [(n, its) for n, its in thu if n not in excluded]
    tue_dict, thu_dict = dict(tue), dict(thu)

    # 代理報告：列管項目掛在owner名下、但由別人(如PM_J)上台報告，config格式 {"PM_J": {"owner": "PM_I", "keyword": "機關己"}}
    proxies = cfg.get('upper_proxies') or {}

    def find_listed_items(name):
        """一位PM這次可能有多個列管項目，全部收集回傳（不只第一個）"""
        if name in proxies:
            p = proxies[name]
            return [it for it in thu_dict.get(p['owner'], []) + tue_dict.get(p['owner'], [])
                    if it['listed'] and p['keyword'] in it['raw']]
        proxied_kw = [p['keyword'] for p in proxies.values() if p['owner'] == name]
        return [it for it in thu_dict.get(name, []) + tue_dict.get(name, [])
                if it['listed'] and not any(k in it['raw'] for k in proxied_kw)]

    upper_items = []
    for name in cfg['upper_order']:
        its = find_listed_items(name)
        if not its:
            raise RuntimeError(f"找不到 {name} 的列管項目，檢查upper_order或listed_projects設定")
        upper_items.append((name, its))
    listed_raw_texts = {it['raw'] for _, its in upper_items for it in its}
    upper_names = {name for name, _ in upper_items}

    # 有列管專案的人整個被拉到週四上半場：週二不再列出他/她（不論原本是否週二PM），
    # 週四下半場也不再重複列出他/她的其他非列管專案（上半場報過就不用再報）。
    lower_groups = []
    for name, items in thu:
        if name in upper_names:
            continue
        remain = [it for it in items if it['raw'] not in listed_raw_texts]
        if remain:
            lower_groups.append((name, remain))

    # ---------- 缺席者處理（排序＋註記）----------
    absentees = cfg.get('absentees', [])
    tue_order = cfg.get('tue_order_override') or [n for n, _ in tue if n not in upper_names]
    tue_order, tue_notes = apply_absentees(tue_order, tue_dict, 'tue', absentees)

    # 週二出差且標記 moved_to=thu 的人，把他/她週二的完整項目整組搬進週四下半場清單
    lower_existing_names = {n for n, _ in lower_groups}
    for a in absentees:
        if a.get('day') == 'tue' and a.get('moved_to') == 'thu':
            name = a['name']
            if name in tue_dict and name not in lower_existing_names:
                lower_groups.append((name, tue_dict[name]))
                lower_existing_names.add(name)

    # 週二會議取消（config `tue_cancelled: true`，2026-11-03新增）：週二PM全部併進週四下半場、排在週四PM前面，不產週二分頁
    tue_cancelled = bool(cfg.get('tue_cancelled'))
    if tue_cancelled:
        moved = [(n, tue_dict[n]) for n in tue_order if n not in lower_existing_names]
        lower_groups = moved + lower_groups
        lower_existing_names.update(n for n, _ in moved)
        tue_order, tue_notes = [], {}

    # 總表沒有、但要列在下半場的人（config `lower_extra`），格式 [{"name": "PM_A", "items": ["機關辛_財產系統建置案：PM_J、成員A", ...]}]
    for ex in cfg.get('lower_extra') or []:
        if ex['name'] not in lower_existing_names:
            lower_groups.append((ex['name'], [{'raw': f"({i}){t}", 'type': None, 'goal': None, 'listed': False}
                                              for i, t in enumerate(ex['items'], 1)]))
            lower_existing_names.add(ex['name'])

    lower_order_default = [n for n, _ in lower_groups]
    lower_order = cfg.get('lower_order_override') or lower_order_default
    lower_order, thu_notes = apply_absentees(lower_order, dict(lower_groups), 'thu', absentees)
    lower_dict = dict(lower_groups)

    # 上半場缺席者同樣排到最後、加註記、不佔時段
    upper_order, upper_notes = apply_absentees([n for n, _ in upper_items], dict(upper_items), 'thu', absentees)
    upper_dict = dict(upper_items)
    # 上半場轉貼分頁的專案文字可用config覆寫（總表用語跟主管會議列管名稱常不同），格式 {"PM_B": ["機關乙_...：成員B、成員C", ...]}
    upper_text = cfg.get('upper_text_override') or {}

    def name_label(name, notes):
        note = notes.get(name)
        return f"{name}{note}" if note else name

    # ==================================================================
    # Stage B: 重寫「專案議程」分頁（含底色/框線）
    # ==================================================================
    for row in agenda.iter_rows(min_row=1, max_row=agenda.max_row, max_col=20):
        for cell in row:
            clear_fill_border_font(cell)

    r = 1
    agenda.cell(row=r, column=1, value='週二報告').font = FONT_TITLE
    r += 1
    agenda.cell(row=r, column=1, value='報告人、專案').font = FONT_HEADER
    agenda.cell(row=r, column=2, value='專案類型').font = FONT_HEADER
    agenda.cell(row=r, column=3, value='專案目標').font = FONT_HEADER
    style_row(agenda, r)
    r += 1
    for name in tue_order:
        agenda.cell(row=r, column=1, value=name_label(name, tue_notes)).font = FONT_NAME
        style_row(agenda, r, is_name=True)
        r += 1
        for it in tue_dict[name]:
            agenda.cell(row=r, column=1, value=it['raw']).font = FONT_PROJ
            agenda.cell(row=r, column=2, value=it['type']).font = FONT_PROJ
            agenda.cell(row=r, column=3, value=it['goal']).font = FONT_PROJ
            style_row(agenda, r)
            if it['listed']:
                agenda.cell(row=r, column=4, value='列管').font = FONT_LISTED
            r += 1
        style_row(agenda, r)
        r += 1

    r += 1
    agenda.cell(row=r, column=1, value='週四報告(上半場主管)').font = FONT_TITLE
    r += 2
    for name, its in upper_items:
        agenda.cell(row=r, column=1, value=name).font = FONT_NAME
        style_row(agenda, r, is_name=True)
        r += 1
        for it in its:
            agenda.cell(row=r, column=1, value=strip_num(it['raw'])).font = FONT_PROJ
            agenda.cell(row=r, column=2, value=it['type']).font = FONT_PROJ
            agenda.cell(row=r, column=3, value=it['goal']).font = FONT_PROJ
            style_row(agenda, r)
            agenda.cell(row=r, column=4, value='列管').font = FONT_LISTED
            r += 1
        style_row(agenda, r)
        r += 1

    r += 1
    agenda.cell(row=r, column=1, value='週四報告(下半場正常)').font = FONT_TITLE
    r += 2
    for name in lower_order:
        agenda.cell(row=r, column=1, value=name_label(name, thu_notes)).font = FONT_NAME
        style_row(agenda, r, is_name=True)
        r += 1
        for idx, it in enumerate(lower_dict[name], 1):
            text = f"({idx}){strip_num(it['raw'])}"
            agenda.cell(row=r, column=1, value=text).font = FONT_PROJ
            agenda.cell(row=r, column=2, value=it['type']).font = FONT_PROJ
            agenda.cell(row=r, column=3, value=it['goal']).font = FONT_PROJ
            style_row(agenda, r)
            r += 1
        style_row(agenda, r)
        r += 1
    print(f"[StageB] 專案議程重寫完成，共 {r} 列")

    # ==================================================================
    # Stage C: 主席／值日生（優先自動查輪值表，沒有這個分頁才用config手動值）
    # ==================================================================
    if rota is not None:
        r_tue, r_thu = find_rota_row(rota, tue_date), find_rota_row(rota, thu_date)
        if r_tue is None or r_thu is None:
            raise RuntimeError("輪值表找不到對應日期，檢查tue_date/thu_date是否已建列")
        chair = rota.cell(row=r_tue, column=6).value
        duty_tue = rota.cell(row=r_tue, column=4).value
        duty_thu = rota.cell(row=r_thu, column=4).value
        boss_chairs = bool(rota.cell(row=r_thu, column=7).value)
        print(f"[StageC] (輪值表自動查得) 主席={chair}　值日生(週二)={duty_tue}　值日生(週四)={duty_thu}　主管與會={boss_chairs}")
    else:
        chair = cfg['chair']
        duty_tue = cfg['duty_tue']
        duty_thu = cfg['duty_thu']
        boss_chairs = bool(cfg['boss_chairs_thu'])
        print(f"[StageC] (config手動值，無輪值表分頁) 主席={chair}　值日生(週二)={duty_tue}　值日生(週四)={duty_thu}　主管與會={boss_chairs}")

    # ==================================================================
    # Stage D: 建立週二／週四轉貼分頁（不含備註/連結，使用者自行補）
    # ==================================================================
    def roc_date(d):
        return f"{d.year - 1911}-{d.month:02d}-{d.day:02d}"

    def build_thu_sheet(title_name):
        ws = wb.create_sheet(title_name)
        r = 1
        c = ws.cell(row=r, column=1, value=f'【{roc_date(thu_date)}會議議程】')
        c.font = Font(name='Microsoft JhengHei', size=12, bold=False)
        r += 1
        chair_line = f'主席：上半場-主管、下半場-{chair}' if boss_chairs else f'主席：{chair}'
        r = write_line(ws, r, chair_line)
        r = write_line(ws, r, f'值日生：{duty_thu}')
        absent_today = [n for n in upper_order if upper_notes.get(n)] + \
                       [n for n in lower_order if thu_notes.get(n)]
        if absent_today:
            r = write_line(ws, r, '請假：' + '、'.join(absent_today))
        r = write_line(ws, r, '-' * 60, bold=True)

        # 每個列管專案10分鐘（2026-10-27使用者定案：兩項的PM排20分鐘），缺席者不佔時段
        def n_items(name):
            return len(upper_text.get(name) or upper_dict[name])
        present = [n for n in upper_order if not upper_notes.get(n)]
        end_min = 13 * 60 + 30 + sum(n_items(n) for n in present) * MIN_PER_ITEM
        r = write_line(ws, r, f"13 : 30 - {end_min // 60:02d} : {end_min % 60:02d} 列管專案報告")
        start2 = end_min + 10
        end2 = start2 + 60
        r = write_line(ws, r, f"{start2 // 60:02d} : {start2 % 60:02d} - {end2 // 60:02d} : {end2 % 60:02d} PM會議")
        r += 1

        t = 13 * 60 + 30
        for name in upper_order:
            if upper_notes.get(name):
                r = write_line(ws, r, name_label(name, upper_notes))
            else:
                t_end = t + n_items(name) * MIN_PER_ITEM
                r = write_line(ws, r, f"{name} {t // 60}:{t % 60:02d}-{t_end // 60}:{t_end % 60:02d}")
                t = t_end
            texts = upper_text.get(name) or [strip_num(it['raw']) for it in upper_dict[name]]
            for idx, text in enumerate(texts, 1):
                r = write_line(ws, r, f"({idx}){text}", bold=False)
            r += 1
        r = write_line(ws, r, '中場休息')
        r += 1
        for name in lower_order:
            r = write_line(ws, r, name_label(name, thu_notes))
            for idx, it in enumerate(lower_dict[name], 1):
                r = write_line(ws, r, f"({idx}){strip_num(it['raw'])}", bold=False)
            r += 1
        return ws, r

    def build_tue_sheet(title_name):
        ws = wb.create_sheet(title_name)
        r = 1
        c = ws.cell(row=r, column=1, value=f'【{roc_date(tue_date)}會議議程】')
        c.font = Font(name='Microsoft JhengHei', size=12, bold=False)
        r += 1
        r = write_line(ws, r, f'主席：{chair}')
        r = write_line(ws, r, f'值日生：{duty_tue}')
        absent_today = [name_label(n, tue_notes) for n in tue_order if tue_notes.get(n)]
        if absent_today:
            r = write_line(ws, r, '請假：' + '、'.join(a.split('(')[0] for a in absent_today))
        r = write_line(ws, r, '-' * 60, bold=True)
        r += 1
        for name in tue_order:
            r = write_line(ws, r, name_label(name, tue_notes))
            for idx, it in enumerate(tue_dict[name], 1):
                r = write_line(ws, r, f"({idx}){strip_num(it['raw'])}", bold=False)
            r += 1
        return ws, r

    thu_title = f"{thu_date.year - 1911:03d}{thu_date.month:02d}{thu_date.day:02d}"
    tue_title = f"{tue_date.year - 1911:03d}{tue_date.month:02d}{tue_date.day:02d}"
    for t in (thu_title, tue_title):
        if t in wb.sheetnames:
            del wb[t]
    ws_thu, n1 = build_thu_sheet(thu_title)
    if tue_cancelled:
        print(f"[StageD] {thu_title} 共{n1}列；週二取消，不產{tue_title}分頁（不含備註/連結區塊）")
    else:
        ws_tue, n2 = build_tue_sheet(tue_title)
        print(f"[StageD] {thu_title} 共{n1}列、{tue_title} 共{n2}列（皆不含備註/連結區塊）")

    # ==================================================================
    # Stage D2: 刪除超過一個月的歷史轉貼分頁，只保留最近期間的分頁
    # （2026-09-14使用者要求：新分頁埋在150+個舊分頁最後面找不到，改為每次跑完順手清舊分頁）
    # ==================================================================
    PROTECTED_SHEETS = {
        '重點專案', '專案總表(正常報告版)', '專案議程', '「115年(會議輪值表)」',
        'PM筆電或其他設備使用登記',
    }
    date_pat = re.compile(r'^\D*(\d{7})')
    cutoff = thu_date - dt.timedelta(days=30)
    removed = []
    for name in list(wb.sheetnames):
        if name in PROTECTED_SHEETS or name in (thu_title, tue_title):
            continue
        m = date_pat.match(name)
        if not m:
            continue
        digits = m.group(1)
        try:
            sheet_date = dt.date(int(digits[0:3]) + 1911, int(digits[3:5]), int(digits[5:7]))
        except ValueError:
            continue
        if sheet_date < cutoff:
            del wb[name]
            removed.append(name)
    if removed:
        print(f"[StageD2] 已刪除 {len(removed)} 個超過一個月({cutoff.isoformat()}前)的舊分頁")
    else:
        print("[StageD2] 沒有超過一個月的舊分頁需要刪除")

    # ---------- Stage D3: 分頁排序，最新一週緊接在固定分頁後面，最左邊最新、最右邊最舊 ----------
    def move_after(name, anchor_name):
        cur = wb.sheetnames.index(name)
        target = wb.sheetnames.index(anchor_name) + 1
        wb.move_sheet(name, offset=target - cur)

    if not tue_cancelled:
        move_after(tue_title, '專案議程')
    move_after(thu_title, '專案議程')
    print(f"[StageD3] 分頁順序已調整：{thu_title}{'' if tue_cancelled else '、' + tue_title} 移到最前面（最新日期在最左）")

    # ==================================================================
    # Stage E（選用）: 重排未來值日生輪值
    # ==================================================================
    dr = cfg.get('duty_rotation')
    if dr and rota is None:
        raise RuntimeError("要重排值日生輪值(duty_rotation)必須提供「115年(會議輪值表)」分頁，才有地方寫入更新結果")
    if dr:
        cycle = dr['cycle']
        start = dt.datetime.fromisoformat(dr['start_date'])
        end = dt.datetime.fromisoformat(dr['end_date'])
        target_rows = []
        for r_ in range(1, rota.max_row + 1):
            a = rota.cell(row=r_, column=1).value
            b = rota.cell(row=r_, column=2).value
            if isinstance(a, dt.datetime) and start <= a <= end and b == dr['weekday']:
                target_rows.append(r_)
        assignments = {r_: cycle[i % len(cycle)] for i, r_ in enumerate(target_rows)}
        # 同週撞名檢查（跟前一列，通常是同週另一天）
        for r_ in target_rows:
            other_name = rota.cell(row=r_ - 1, column=4).value
            if assignments[r_] == other_name:
                idx = target_rows.index(r_)
                if idx + 1 < len(target_rows):
                    nxt = target_rows[idx + 1]
                    assignments[r_], assignments[nxt] = assignments[nxt], assignments[r_]
                    print(f"  [撞名修正] {rota.cell(row=r_, column=1).value:%Y-%m-%d} 與 {other_name} 撞名，"
                          f"與 {rota.cell(row=nxt, column=1).value:%Y-%m-%d} 互換")
        style_ref = rota.cell(row=target_rows[0] - 1, column=4).font if target_rows else Font()
        for r_ in target_rows:
            cell = rota.cell(row=r_, column=4, value=assignments[r_])
            cell.font = copy(style_ref)
        print(f"[StageE] 輪值表 {dr['weekday']} 已更新 {len(target_rows)} 筆（{dr['start_date']}~{dr['end_date']}）")

    wb.save(cfg['output_xlsx'])
    print(f"\n✓ 已輸出：{cfg['output_xlsx']}")


if __name__ == '__main__':
    main(sys.argv[1])
