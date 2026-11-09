# -*- coding: utf-8 -*-
"""
================================================================================
나이스(NEIS) 출결 자동 입력 프로그램 (v2.0)
================================================================================
- 학교 담임교사를 위한 나이스 일일 출결 자동 입력 RPA 프로그램
- 변경된 새 엑셀 양식 [일자 | 학번 | 교시 | 출결구분 | 상태분류 | 시스템등록일시] 지원
- 안전 규칙:
  1) 키보드 보안 프로그램 충돌(BSOD) 방지를 위한 자바스크립트 값 직접 설정
  2) 화면 제목 날짜와 목표 날짜 일치 여부 철저 검증
  3) 교시칸 1회 클릭 및 '/' 상태 확인 (재클릭 절대 금지)
  4) 테스트 모드 지원 (팝업 닫기 및 교시칸 미클릭)
  5) 브라우저 전용 프로필 사용 및 connect_over_cdp 연결
================================================================================
"""

import os
import sys
import re
import csv
import time
import json
import queue
import datetime
import threading
import subprocess
import traceback
from typing import Dict, List, Optional, Tuple, Any

# GUI 라이브러리
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# 엑셀 처리
import openpyxl

# 브라우저 자동화
from playwright.sync_api import sync_playwright, Page, Frame, Dialog

# ----------------------------------------------------------------------
# 상수 및 전역 설정
# ----------------------------------------------------------------------
APP_NAME = "나이스 출결 자동 입력 프로그램"
APP_VERSION = "v2.0 (새 엑셀 양식)"
DEFAULT_NEIS_URL = "https://gne.eduptl.kr/bpm_lgn_lg00_001.do"
DEFAULT_CDP_PORT = 9222

# 나이스 교시 순서 가중치 (정렬용)
PERIOD_ORDER = {
    '조례': 0,
    '조회': 0,
    '1교시': 1,
    '2교시': 2,
    '3교시': 3,
    '4교시': 4,
    '5교시': 5,
    '6교시': 6,
    '7교시': 7,
    '8교시': 8,
    '종례': 99,
}

# ----------------------------------------------------------------------
# 유틸리티 함수: 날짜, 학번, 교시 정규화 및 출결 해석
# ----------------------------------------------------------------------
def normalize_date(val: Any) -> str:
    """날짜 값을 'YYYY.MM.DD' 형식의 표준 문자열로 변환합니다."""
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%Y.%m.%d")
    s = str(val).strip().replace("-", ".").replace("/", ".")
    parts = [p for p in s.split(".") if p]
    if len(parts) >= 3:
        y = parts[0]
        m = parts[1].zfill(2)
        d = parts[2].zfill(2)
        return f"{y}.{m}.{d}"
    return s

def parse_student_num(val: Any) -> Tuple[Optional[int], str]:
    """
    학번 및 학년/반 파싱:
    - 17, 9 등 1~2자리: 번호=17, 9 / 학년반=""
    - 20917 등 3자리 이상: 번호=17 (마지막 2자리), 학년반="209"
    """
    if val is None:
        return None, ""
    s = str(val).strip()
    digits = "".join([c for c in s if c.isdigit()])
    if not digits:
        return None, ""
    if len(digits) >= 3:
        num = int(digits[-2:])
        grade_class = digits[:-2]
        return num, grade_class
    else:
        return int(digits), ""

def normalize_period(val: Any) -> str:
    """교시 표기를 나이스 표준('조회', '1교시', ..., '종례')으로 변환합니다."""
    if val is None:
        return ""
    s = str(val).strip()
    if s == '조례':
        return '조회'
    if s.isdigit():
        return f"{s}교시"
    return s

def get_period_sort_key(period_str: str) -> int:
    """교시 정렬을 위한 가중치 반환"""
    p = str(period_str).strip()
    if p in PERIOD_ORDER:
        return PERIOD_ORDER[p]
    m = re.search(r'\d+', p)
    if m:
        return int(m.group())
    return 50

def parse_status(status_raw: Any, category_raw: Any = "") -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    출결구분 및 상태분류 해석:
    1) 구분(앞부분):
       - '미인정' 또는 '무단' -> '미인정'
       - '인정' -> '출석인정'
       - '병' -> '질병'
       - '기타' -> '기타'
       (※ '미인정'을 '인정'보다 먼저 판별해야 함)
    2) 종류(뒷부분):
       - '결과'로 끝남 -> '결과' (※ '결과'를 '결'보다 먼저 판별해야 함)
       - '지각'으로 끝남 -> '지각'
       - '조퇴'로 끝남 -> '조퇴'
       - '결'로 끝남 (예: 인정결, 병결, 미인정결) -> '결석'
    3) 상태분류 검증:
       - 결석류/지각류/조퇴류/결과류 와 파싱된 종류가 불일치하면 엑셀오류
    """
    s = str(status_raw).strip() if status_raw else ""
    c = str(category_raw).strip() if category_raw else ""

    if not s:
        return None, None, "출결구분이 비어 있습니다."

    # 1. 구분 판별 ('미인정' 우선)
    division = None
    if '미인정' in s or '무단' in s:
        division = '미인정'
    elif '인정' in s:
        division = '출석인정'
    elif '병' in s:
        division = '질병'
    elif '기타' in s:
        division = '기타'

    # 2. 종류 판별 ('결과' 우선)
    kind = None
    if s.endswith('결과'):
        kind = '결과'
    elif s.endswith('지각'):
        kind = '지각'
    elif s.endswith('조퇴'):
        kind = '조퇴'
    elif s.endswith('결'):
        kind = '결석'

    if not division:
        return None, kind, f"인식할 수 없는 출결구분: '{s}'"
    if not kind:
        return division, None, f"인식할 수 없는 출결종류: '{s}'"

    # 3. 상태분류 일치 여부 확인
    category_map = {
        '결석류': '결석',
        '지각류': '지각',
        '조퇴류': '조퇴',
        '결과류': '결과'
    }
    if c:
        for cat_key, exp_kind in category_map.items():
            if cat_key in c:
                if kind != exp_kind:
                    return division, kind, f"출결구분({s})과 상태분류({c}) 불일치"
                break

    return division, kind, None

# ----------------------------------------------------------------------
# 엑셀 파일 읽기 및 학생별 기록 묶기
# ----------------------------------------------------------------------
def read_attendance_excel(filepath: str) -> Dict[str, Any]:
    """
    엑셀 파일을 읽고, 학생/일자/출결구분 별로 묶어 정렬된 기록 구조를 생성합니다.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {filepath}")

    # 잠금 파일(~$파일명) 검사: 엑셀이 열려 있는 경우 감지
    dir_name = os.path.dirname(os.path.abspath(filepath))
    base_name = os.path.basename(filepath)
    lock_file = os.path.join(dir_name, f"~${base_name}")
    if os.path.exists(lock_file):
        raise IOError(f"엑셀 파일이 현재 열려 있습니다 ('~${base_name}' 감지).\n엑셀을 저장(Ctrl+S) 후 닫고 다시 실행해 주세요.")

    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet = wb.active

    # 병합된 셀(merged cells)의 값을 왼쪽 위 셀 값으로 사전 매핑
    merged_map = {}
    for rng in sheet.merged_cells.ranges:
        top_left_val = sheet.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                merged_map[(r, col)] = top_left_val

    def get_cell(r, c):
        if (r, c) in merged_map:
            return merged_map[(r, c)]
        return sheet.cell(r, c).value

    # 1. 머리글 행 자동 탐색
    header_aliases = {
        'date': ['일자', '날짜', 'date'],
        'num': ['학번', '번호', 'student_no', 'id'],
        'period': ['교시', '교시구분', 'period'],
        'status': ['출결구분', '출결', '구분', 'status'],
        'category': ['상태분류', '분류', 'category'],
        'name': ['이름', '성명', 'name'],
        'reason': ['사유', '비고', 'reason'],
    }

    header_row = None
    col_map = {}
    max_search_row = min(sheet.max_row, 30)

    for r in range(1, max_search_row + 1):
        row_vals = {}
        for c in range(1, sheet.max_column + 1):
            val = get_cell(r, c)
            if val is not None:
                row_vals[c] = str(val).strip()

        temp_map = {}
        for field, aliases in header_aliases.items():
            for c, val_str in row_vals.items():
                if any(alias in val_str for alias in aliases):
                    if field not in temp_map:
                        temp_map[field] = c
                        break

        # 필수 열 ('일자', '학번', '교시', '출결구분') 확인
        if 'date' in temp_map and 'num' in temp_map and 'period' in temp_map and 'status' in temp_map:
            header_row = r
            col_map = temp_map
            break

    if not header_row:
        raise ValueError("엑셀에서 필수 머리글('일자', '학번', '교시', '출결구분')을 찾을 수 없습니다.")

    # 2. 데이터 행 순회 및 파싱
    raw_rows_count = 0
    grouped_dict = {}  # key: (date, num, division, kind)
    error_list = []

    for r in range(header_row + 1, sheet.max_row + 1):
        raw_date = get_cell(r, col_map['date'])
        raw_num = get_cell(r, col_map['num'])
        raw_period = get_cell(r, col_map['period'])
        raw_status = get_cell(r, col_map['status'])
        raw_cat = get_cell(r, col_map['category']) if 'category' in col_map else ""

        # 빈 행 건너뛰기
        if raw_date is None and raw_num is None and raw_period is None and raw_status is None:
            continue

        raw_rows_count += 1

        norm_date = normalize_date(raw_date)
        num, grade_class = parse_student_num(raw_num)
        period = normalize_period(raw_period)
        division, kind, err = parse_status(raw_status, raw_cat)

        if not norm_date or num is None or not period or err:
            err_msg = err or "필수 값 누락 (일자/학번/교시)"
            error_list.append({
                'row': r,
                'date': norm_date or str(raw_date),
                'num': num if num is not None else str(raw_num),
                'period': period or str(raw_period),
                'status': str(raw_status),
                'category': str(raw_cat),
                'error': err_msg
            })
            continue

        group_key = (norm_date, num, division, kind)
        if group_key not in grouped_dict:
            grouped_dict[group_key] = {
                'date': norm_date,
                'student_num': num,
                'student_grade_class': grade_class,
                'division': division,
                'kind': kind,
                'periods': [],
                'source_rows': [],
                'status_raw': str(raw_status),
                'category_raw': str(raw_cat),
                'target_period': None,
            }

        rec = grouped_dict[group_key]
        if period not in rec['periods']:
            rec['periods'].append(period)
        rec['source_rows'].append(r)

    # 3. 교시 정렬 및 클릭 대상 교시 선정
    records_by_date = {}
    for group_key, rec in grouped_dict.items():
        # 나이스 열 순서대로 교시 목록 정렬 (조회 -> 1교시 ... 7교시 -> 종례)
        rec['periods'].sort(key=get_period_sort_key)

        # 클릭 대상 교시 결정:
        # - 지각: 목록 중 '마지막' 교시 (예: [조회, 1교시] -> 1교시)
        # - 조퇴: 목록 중 '첫' 교시 (예: [3교시, 4교시, 5교시, 6교시] -> 3교시)
        # - 결석: 교시칸 클릭 없음 (마감칸만 처리)
        # - 결과: 교시 목록의 각 교시칸 개별 팝업 처리
        if rec['kind'] == '지각':
            rec['target_period'] = rec['periods'][-1] if rec['periods'] else None
        elif rec['kind'] == '조퇴':
            rec['target_period'] = rec['periods'][0] if rec['periods'] else None
        elif rec['kind'] == '결석':
            rec['target_period'] = None
        elif rec['kind'] == '결과':
            rec['target_period'] = None

        d = rec['date']
        if d not in records_by_date:
            records_by_date[d] = []
        records_by_date[d].append(rec)

    # 날짜 오름차순 및 날짜 내 학번 오름차순 정렬
    sorted_records_by_date = {}
    for d in sorted(records_by_date.keys()):
        records_by_date[d].sort(key=lambda x: x['student_num'])
        sorted_records_by_date[d] = records_by_date[d]

    total_records = sum(len(recs) for recs in sorted_records_by_date.values())

    return {
        'header_row': header_row,
        'col_map': col_map,
        'records_by_date': sorted_records_by_date,
        'total_rows': raw_rows_count,
        'total_records': total_records,
        'errors': error_list
    }

# ----------------------------------------------------------------------
# 로깅 및 결과 기록기 (run.log 즉시 동기화, 처리결과.csv)
# ----------------------------------------------------------------------
class RpaLogger:
    """로그 파일, CSV 결과 파일, 스크린샷 저장을 관리하는 클래스"""
    def __init__(self, log_dir: str):
        self.log_dir = log_dir
        os.makedirs(self.log_dir, exist_ok=True)
        self.log_file_path = os.path.join(self.log_dir, "run.log")
        self.csv_file_path = os.path.join(self.log_dir, "처리결과.csv")
        self.log_file = open(self.log_file_path, "a", encoding="utf-8")
        self.ui_callback = None

        # CSV 파일 초기화
        self._init_csv()

    def _init_csv(self):
        if not os.path.exists(self.csv_file_path) or os.path.getsize(self.csv_file_path) == 0:
            with open(self.csv_file_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["일자", "번호", "구분", "종류", "교시목록", "원래엑셀행번호", "결과", "메시지"])

    def set_ui_callback(self, cb):
        self.ui_callback = cb

    def log(self, message: str, level: str = "INFO"):
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{now_str}] [{level}] {message}\n"
        self.log_file.write(line)
        # PC가 갑자기 꺼져도 마지막 단계가 디스크에 물리적으로 남도록 즉시 동기화
        self.log_file.flush()
        try:
            os.fsync(self.log_file.fileno())
        except Exception:
            pass

        if self.ui_callback:
            self.ui_callback(message, level)

    def log_step(self, step_name: str):
        self.log(f"[단계] {step_name}", level="STEP")

    def record_csv(self, date: str, num: int, division: str, kind: str,
                   periods: List[str], source_rows: List[int],
                   result_status: str, message: str = ""):
        """처리 결과를 CSV 파일에 1줄 추가합니다."""
        with open(self.csv_file_path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([
                date,
                num,
                division,
                kind,
                ",".join(periods),
                ",".join(map(str, source_rows)),
                result_status,
                message
            ])

    def save_screenshot(self, page: Page, name_prefix: str) -> str:
        """현재 화면의 스크린샷을 저장합니다."""
        now_str = datetime.datetime.now().strftime("%H%M%S_%f")[:9]
        filename = f"{name_prefix}_{now_str}.png"
        filepath = os.path.join(self.log_dir, filename)
        try:
            page.screenshot(path=filepath, full_page=True)
            self.log(f"스크린샷 저장 완료: {filepath}", level="DEBUG")
            return filepath
        except Exception as e:
            self.log(f"스크린샷 저장 실패: {e}", level="WARN")
            return ""

    def close(self):
        try:
            self.log_file.close()
        except Exception:
            pass

# ----------------------------------------------------------------------
# 브라우저 실행 및 제어 관련 헬퍼
# ----------------------------------------------------------------------
def get_browser_exe_path(browser_choice: str) -> str:
    """Chrome 또는 Edge 실행 파일 경로를 탐색합니다."""
    candidates = []
    if browser_choice.lower() == 'chrome':
        candidates = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")
        ]
    else:
        candidates = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe")
        ]

    for p in candidates:
        if os.path.exists(p):
            return p
    return ""

def launch_remote_browser(browser_choice: str, port: int, target_url: str) -> bool:
    """
    원격 제어 포트(--remote-debugging-port)와 전용 프로필 폴더로 브라우저를 실행합니다.
    (최신 크롬/엣지는 기본 프로필로는 원격 제어를 차단하므로 전용 프로필 폴더 필수)
    """
    exe_path = get_browser_exe_path(browser_choice)
    if not exe_path:
        raise FileNotFoundError(f"{browser_choice} 실행 파일을 찾을 수 없습니다. 설치 여부를 확인해 주세요.")

    local_app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    profile_sub = "chrome_profile" if browser_choice.lower() == 'chrome' else "edge_profile"
    profile_dir = os.path.join(local_app_data, "NeisAttendanceRPA", profile_sub)
    os.makedirs(profile_dir, exist_ok=True)

    args = [
        exe_path,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile_dir}",
        target_url
    ]

    creationflags = 0
    if sys.platform == "win32":
        creationflags = subprocess.DETACHED_PROCESS

    subprocess.Popen(args, creationflags=creationflags)
    return True

# ----------------------------------------------------------------------
# 자바스크립트 헬퍼 코드 (안전한 값 설정, DOM 탐색, 테이블 열 계산)
# ----------------------------------------------------------------------
# 키보드 보안 프로그램 블루스크린 방지: 네이티브 value setter + 이벤트 디스패치
SAFE_SET_VALUE_JS = """
(element, value) => {
    if (!element) return;
    const proto = (element instanceof HTMLTextAreaElement) ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
    if (setter) {
        setter.call(element, value);
    } else {
        element.value = value;
    }
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
    element.dispatchEvent(new Event('blur', { bubbles: true }));
}
"""

# 화면 상단 제목에서 현재 날짜 추출 ("2학년 2반 2026.11.09. Total 26" 등)
GET_SCREEN_DATE_JS = """
() => {
    // 1. 특정 제목 요소 탐색
    const allElems = Array.from(document.querySelectorAll('span, div, h1, h2, h3, h4, p, label'));
    for (const el of allElems) {
        const text = el.innerText ? el.innerText.trim() : '';
        // "2026.11.09." 또는 "2026.11.09" 패턴 검색
        const m = text.match(/(\\d{4})\\.(\\d{2})\\.(\\d{2})/);
        if (m && (text.includes('반') || text.includes('Total') || text.includes('학년'))) {
            return `${m[1]}.${m[2]}.${m[3]}`;
        }
    }
    // 2. 일자 입력칸 값 확인
    const inputs = Array.from(document.querySelectorAll('input[type="text"], input:not([type])'));
    for (const inp of inputs) {
        const val = (inp.value || '').trim();
        const m = val.match(/(\\d{4})\\.(\\d{2})\\.(\\d{2})/);
        if (m) {
            return `${m[1]}.${m[2]}.${m[3]}`;
        }
    }
    return "";
}
"""

# 명단 표 머리글 줄('번호' 기준 ±25px)에서 열(번호, 성명, 마감, 조회, 1~7교시, 종례) 위치 계산
GET_HEADER_LAYOUT_JS = """
() => {
    // 1. '번호' 머리글 셀 찾기
    const cells = Array.from(document.querySelectorAll('th, td, div, span'));
    let numCell = null;
    for (const c of cells) {
        const t = (c.innerText || '').trim();
        if (t === '번호') {
            const rect = c.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0 && rect.top > 50) {
                numCell = c;
                break;
            }
        }
    }
    if (!numCell) return null;

    const numRect = numCell.getBoundingClientRect();
    const targetY = numRect.top + numRect.height / 2;

    // 2. targetY 기준 ±25px 범위에 있는 머리글 요소들 수집
    // 상단 [조회] 버튼과 헷갈리지 않도록 반드시 명단 머리글 높이에 있는 것만 채택
    const headerRowElems = cells.filter(c => {
        const r = c.getBoundingClientRect();
        return r.width > 0 && r.height > 0 && Math.abs((r.top + r.height / 2) - targetY) <= 25;
    });

    const layout = {};
    for (const el of headerRowElems) {
        const t = (el.innerText || '').trim();
        const r = el.getBoundingClientRect();
        if (!t) continue;

        if (t === '번호' && !layout['번호']) {
            layout['번호'] = { left: r.left, right: r.right, text: t };
        } else if ((t === '성명' || t === '이름') && !layout['성명']) {
            layout['성명'] = { left: r.left, right: r.right, text: t };
        } else if (t === '마감' && !layout['마감']) {
            layout['마감'] = { left: r.left, right: r.right, text: t };
        } else if (t === '조회' && !layout['조회']) {
            layout['조회'] = { left: r.left, right: r.right, text: t };
        } else if (t === '종례' && !layout['종례']) {
            layout['종례'] = { left: r.left, right: r.right, text: t };
        } else if (t === '사유' && !layout['사유']) {
            layout['사유'] = { left: r.left, right: r.right, text: t };
        } else {
            // N교시(과목) 형태 확인
            const m = t.match(/(\\d+)교시/);
            if (m) {
                const key = `${m[1]}교시`;
                if (!layout[key]) {
                    layout[key] = { left: r.left, right: r.right, text: t };
                }
            }
        }
    }
    return { numRect: { top: numRect.top, bottom: numRect.bottom }, layout: layout };
}
"""

# 학생 행 및 칸 탐색, 가림 및 완전 노출 검사
FIND_STUDENT_CELL_JS = """
(args) => {
    const { studentNum, colName } = args;

    // 1. 표 스크롤 컨테이너 찾기
    const tables = Array.from(document.querySelectorAll('table, tbody, div[class*="grid"], div[class*="scroll"]'));
    let scrollContainer = null;
    for (const t of tables) {
        if (t.scrollHeight > t.clientHeight && t.clientHeight > 100) {
            scrollContainer = t;
            break;
        }
    }
    if (!scrollContainer) {
        scrollContainer = document.scrollingElement || document.body;
    }

    // 2. 명단 머리글 줄 열 레이아웃 다시 확인
    const cells = Array.from(document.querySelectorAll('th, td, div, span'));
    let numHeader = null;
    for (const c of cells) {
        if ((c.innerText || '').trim() === '번호') {
            const r = c.getBoundingClientRect();
            if (r.width > 0 && r.height > 0 && r.top > 50) {
                numHeader = c;
                break;
            }
        }
    }
    if (!numHeader) return { status: 'NO_HEADER' };
    const numHRect = numHeader.getBoundingClientRect();
    const hY = numHRect.top + numHRect.height / 2;

    const headerCells = cells.filter(c => {
        const r = c.getBoundingClientRect();
        return r.width > 0 && r.height > 0 && Math.abs((r.top + r.height / 2) - hY) <= 25;
    });

    let targetColLeft = -1, targetColRight = -1;
    let numColLeft = numHRect.left, numColRight = numHRect.right;
    let nameColLeft = -1, nameColRight = -1;

    for (const hc of headerCells) {
        const t = (hc.innerText || '').trim();
        const r = hc.getBoundingClientRect();
        if ((t === '성명' || t === '이름') && nameColLeft === -1) {
            nameColLeft = r.left; nameColRight = r.right;
        }
        if (colName === 'magam' && t === '마감') {
            targetColLeft = r.left; targetColRight = r.right;
        } else if (colName === '조회' && t === '조회') {
            targetColLeft = r.left; targetColRight = r.right;
        } else if (colName === '종례' && t === '종례') {
            targetColLeft = r.left; targetColRight = r.right;
        } else if (t.includes(colName)) {
            targetColLeft = r.left; targetColRight = r.right;
        }
    }

    // 3. 학생 번호 셀 찾기 (번호 열 범위 안에 위치하며 숫자가 정확히 일치하는 셀)
    const allCells = Array.from(document.querySelectorAll('td, div, span'));
    const visibleNums = [];
    let matchedNumCell = null;
    let matchedRow = null;

    for (const el of allCells) {
        const text = (el.innerText || '').trim();
        if (!text || !/^\\d+$/.test(text)) continue;
        const r = el.getBoundingClientRect();
        if (r.width === 0 || r.height === 0) continue;

        // X 좌표가 번호 열 범위와 겹치는지 확인
        const xOverlap = Math.max(0, Math.min(r.right, numColRight) - Math.max(r.left, numColLeft));
        if (xOverlap > 5 && r.top > numHRect.bottom) {
            const val = parseInt(text, 10);
            visibleNums.push(val);
            if (val === studentNum) {
                matchedNumCell = el;
                // 가장 가까운 TR 또는 행 컨테이너 탐색
                matchedRow = el.closest('tr') || el.closest('div[role="row"]') || el.parentElement;
                break;
            }
        }
    }

    if (!matchedNumCell) {
        visibleNums.sort((a, b) => a - b);
        return {
            status: 'NUM_NOT_FOUND',
            visibleNums: visibleNums,
            canScrollDown: scrollContainer.scrollTop + scrollContainer.clientHeight < scrollContainer.scrollHeight - 5,
            canScrollUp: scrollContainer.scrollTop > 5
        };
    }

    // 학생 이름 읽기
    let studentName = "";
    if (matchedRow && nameColLeft !== -1) {
        const rowChildren = Array.from(matchedRow.querySelectorAll('*'));
        for (const child of rowChildren) {
            const cr = child.getBoundingClientRect();
            if (cr.width > 0 && cr.height > 0) {
                const xOver = Math.max(0, Math.min(cr.right, nameColRight) - Math.max(cr.left, nameColLeft));
                if (xOver > 5) {
                    const txt = (child.innerText || '').trim();
                    if (txt && isNaN(txt)) {
                        studentName = txt;
                        break;
                    }
                }
            }
        }
    }

    // 4. 대상 칸(마감 또는 특정 교시) 셀 찾기
    let targetCell = null;
    if (matchedRow && targetColLeft !== -1) {
        const rowDescendants = Array.from(matchedRow.querySelectorAll('td, div, span, button'));
        for (const desc of rowDescendants) {
            const dr = desc.getBoundingClientRect();
            if (dr.width > 0 && dr.height > 0) {
                const xOver = Math.max(0, Math.min(dr.right, targetColRight) - Math.max(dr.left, targetColLeft));
                if (xOver > 5) {
                    targetCell = desc;
                    break;
                }
            }
        }
    }

    if (!targetCell) {
        return {
            status: 'CELL_NOT_FOUND',
            studentName: studentName,
            targetColFound: targetColLeft !== -1
        };
    }

    // 5. 완전 노출 검사 (행 위·아래 끝이 뷰포트 내에 있고 가려지지 않았는지)
    const rowRect = matchedRow.getBoundingClientRect();
    const cellRect = targetCell.getBoundingClientRect();
    const numRect = matchedNumCell.getBoundingClientRect();

    const containerRect = scrollContainer.getBoundingClientRect();
    const isWithinVerticalBounds = (rowRect.top >= containerRect.top - 2) && (rowRect.bottom <= containerRect.bottom + 2);

    // 3개 지점(상단 20%, 중앙 50%, 하단 80%) elementFromPoint 검사
    const pointsToCheck = [
        { x: cellRect.left + cellRect.width / 2, y: cellRect.top + cellRect.height * 0.2 },
        { x: cellRect.left + cellRect.width / 2, y: cellRect.top + cellRect.height * 0.5 },
        { x: cellRect.left + cellRect.width / 2, y: cellRect.top + cellRect.height * 0.8 }
    ];

    let isOccluded = false;
    for (const pt of pointsToCheck) {
        const elAtPoint = document.elementFromPoint(pt.x, pt.y);
        if (!elAtPoint || (!targetCell.contains(elAtPoint) && !elAtPoint.contains(targetCell))) {
            isOccluded = true;
            break;
        }
    }

    const isFullyVisible = isWithinVerticalBounds && !isOccluded;

    return {
        status: 'OK',
        studentName: studentName,
        isFullyVisible: isFullyVisible,
        cellText: (targetCell.innerText || '').trim(),
        cellRect: {
            x: cellRect.left,
            y: cellRect.top,
            width: cellRect.width,
            height: cellRect.height
        },
        numRect: {
            x: numRect.left,
            y: numRect.top,
            width: numRect.width,
            height: numRect.height
        }
    };
}
"""

# 학생 행의 전체 교시(조회~종례) 텍스트 읽기 (클릭 전후 비교용)
GET_ROW_PERIOD_VALUES_JS = """
(studentNum) => {
    // 1. 학생 번호 셀 및 행 찾기
    const cells = Array.from(document.querySelectorAll('td, div, span'));
    let matchedRow = null;
    for (const el of cells) {
        const t = (el.innerText || '').trim();
        if (t === String(studentNum)) {
            const r = el.getBoundingClientRect();
            if (r.width > 0 && r.height > 0 && r.top > 80) {
                matchedRow = el.closest('tr') || el.parentElement;
                break;
            }
        }
    }
    if (!matchedRow) return {};

    // 2. 머리글 위치 확인
    const headerCells = cells.filter(c => {
        const r = c.getBoundingClientRect();
        return r.width > 0 && r.height > 0 && r.top > 40 && r.top < 150;
    });

    const result = {};
    const cols = ['마감', '조회', '1교시', '2교시', '3교시', '4교시', '5교시', '6교시', '7교시', '종례'];

    for (const colName of cols) {
        let colLeft = -1, colRight = -1;
        for (const hc of headerCells) {
            const ht = (hc.innerText || '').trim();
            if (ht.includes(colName)) {
                const hr = hc.getBoundingClientRect();
                colLeft = hr.left;
                colRight = hr.right;
                break;
            }
        }
        if (colLeft !== -1) {
            const rowDescendants = Array.from(matchedRow.querySelectorAll('*'));
            for (const desc of rowDescendants) {
                const dr = desc.getBoundingClientRect();
                if (dr.width > 0 && dr.height > 0) {
                    const xOver = Math.max(0, Math.min(dr.right, colRight) - Math.max(dr.left, colLeft));
                    if (xOver > 5) {
                        result[colName] = (desc.innerText || '').trim();
                        break;
                    }
                }
            }
        }
    }
    return result;
}
"""

# 확인창 버튼 클릭 처리 (제목 "확인"과 버튼 "확인" 엄격 구분)
HANDLE_CONFIRM_MODAL_JS = """
() => {
    // 팝업/모달 컨테이너 탐색
    const dialogs = Array.from(document.querySelectorAll('div[role="dialog"], div[class*="popup"], div[class*="modal"], div[class*="window"], div[class*="alert"]'));
    let activeModal = null;
    for (const d of dialogs) {
        const r = d.getBoundingClientRect();
        if (r.width > 100 && r.height > 60 && window.getComputedStyle(d).display !== 'none' && window.getComputedStyle(d).visibility !== 'hidden') {
            activeModal = d;
            break;
        }
    }

    const scope = activeModal || document;
    // 제목이 아닌 '버튼' 역할을 하는 요소만 탐색
    const candidates = Array.from(scope.querySelectorAll('button, a, input[type="button"], input[type="submit"], [role="button"]'));

    let confirmBtn = null;
    for (const b of candidates) {
        const text = (b.innerText || b.value || '').trim();
        if (text === '확인' || text.includes('확인')) {
            const r = b.getBoundingClientRect();
            // 제목 바(상단)가 아닌 창 아래쪽에 위치한 버튼 우선
            if (r.width > 20 && r.height > 15) {
                confirmBtn = b;
                break;
            }
        }
    }

    if (confirmBtn) {
        confirmBtn.click();
        return { clicked: true, text: confirmBtn.innerText || confirmBtn.value };
    }
    return { clicked: false };
}
"""

# ----------------------------------------------------------------------
# 나이스 RPA 실행 제어 엔진
# ----------------------------------------------------------------------
class NeisAttendanceRPA:
    def __init__(self, logger: RpaLogger, options: Dict[str, Any]):
        self.logger = logger
        self.options = options
        self.stop_requested = False
        self.playwright = None
        self.browser = None
        self.page = None
        self.frame = None

    def request_stop(self):
        """자동화 즉시 중지 요청"""
        self.stop_requested = True
        self.logger.log("사용자에 의해 중지 요청이 수신되었습니다.", level="WARN")

    def check_stop(self):
        """중지 요청 여부 확인"""
        if self.stop_requested:
            raise InterruptedError("사용자가 작업을 중지했습니다.")

    def sleep_safe(self, seconds: float):
        """느리게 실행 옵션 및 중지 플래그를 고려한 안전한 대기"""
        factor = 1.6 if self.options.get('slow_mode', False) else 1.0
        total_time = seconds * factor
        start = time.time()
        while time.time() - start < total_time:
            self.check_stop()
            time.sleep(0.05)

    def connect_browser(self, port: int) -> bool:
        """Playwright를 통해 이미 실행된 브라우저에 CDP로 연결합니다."""
        self.logger.log(f"CDP 연결 시도 중 (http://127.0.0.1:{port})...")
        self.playwright = sync_playwright().start()

        try:
            self.browser = self.playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
        except Exception as e:
            self.logger.log(f"브라우저 CDP 연결 실패: {e}\n브라우저가 9222 포트로 실행 중인지 확인하세요.", level="ERROR")
            return False

        contexts = self.browser.contexts
        if not contexts:
            self.logger.log("브라우저 컨텍스트를 찾을 수 없습니다.", level="ERROR")
            return False

        # 열려 있는 모든 탭 확인
        target_page = None
        for ctx in contexts:
            for p in ctx.pages:
                # 브라우저 기본 alert/confirm 자동 수락 등록
                p.on("dialog", lambda dialog: dialog.accept())
                title = p.title()
                url = p.url
                if "gne.eduptl.kr" in url or "출결" in title or "나이스" in title:
                    target_page = p
                    break
            if target_page:
                break

        if not target_page and contexts[0].pages:
            target_page = contexts[0].pages[0]

        if not target_page:
            self.logger.log("연결할 브라우저 탭(페이지)을 찾을 수 없습니다.", level="ERROR")
            return False

        self.page = target_page
        self.page.bring_to_front()
        self.logger.log(f"브라우저 탭 연결 완료: '{self.page.title()}' ({self.page.url})")
        return True

    def find_active_frame(self) -> Frame:
        """나이스 명단 표가 들어있는 프레임(iframe)을 탐색합니다."""
        for frame in self.page.frames:
            try:
                has_content = frame.evaluate("""() => {
                    const txt = document.body ? document.body.innerText : '';
                    return (txt.includes('출결마감') || txt.includes('마감')) && (txt.includes('번호') && (txt.includes('성명') || txt.includes('이름')));
                }""")
                if has_content:
                    return frame
            except Exception:
                continue
        return self.page.main_frame

    def get_screen_date(self) -> str:
        """화면 상단 제목 또는 일자 입력칸에서 현재 표시된 날짜를 읽어옵니다."""
        frame = self.find_active_frame()
        try:
            val = frame.evaluate(GET_SCREEN_DATE_JS)
            return val or ""
        except Exception as e:
            self.logger.log(f"화면 날짜 읽기 오류: {e}", level="DEBUG")
            return ""

    def click_inquiry_button(self, frame: Frame) -> bool:
        """상단 [조회] 버튼을 클릭합니다."""
        self.logger.log_step("조회 버튼 클릭")
        try:
            clicked = frame.evaluate("""() => {
                const buttons = Array.from(document.querySelectorAll('button, a, input[type="button"], [role="button"]'));
                for (const b of buttons) {
                    const t = (b.innerText || b.value || '').trim();
                    if (t === '조회') {
                        const r = b.getBoundingClientRect();
                        // 상단 조회 버튼은 높이가 보통 150px 이내에 위치
                        if (r.width > 20 && r.height > 15 && r.top < 150) {
                            b.click();
                            return true;
                        }
                    }
                }
                return false;
            }""")
            return clicked
        except Exception as e:
            self.logger.log(f"조회 버튼 클릭 실패: {e}", level="WARN")
            return False

    def wait_table_loaded(self, frame: Frame, max_wait: float = 6.0) -> bool:
        """
        고정 대기 대신 '명단 번호가 화면에 보이고 두 번 연속 같으면' 로딩 완료로 판단합니다.
        """
        start = time.time()
        last_nums = None
        while time.time() - start < max_wait:
            self.check_stop()
            try:
                nums = frame.evaluate("""() => {
                    const cells = Array.from(document.querySelectorAll('td, div, span'));
                    const list = [];
                    for (const c of cells) {
                        const t = (c.innerText || '').trim();
                        if (/^\\d+$/.test(t)) {
                            const r = c.getBoundingClientRect();
                            if (r.width > 0 && r.height > 0 && r.top > 80) {
                                list.push(parseInt(t, 10));
                            }
                        }
                    }
                    return list.slice(0, 10);
                }""")
                if nums and len(nums) > 0:
                    if last_nums == nums:
                        return True
                    last_nums = nums
            except Exception:
                pass
            time.sleep(0.3)
        return False

    def navigate_to_date(self, target_date: str) -> bool:
        """
        목표 날짜로 이동합니다:
        ① 달력 아이콘 클릭 -> 연/월 이동 -> 날짜 클릭 -> [조회]
        ② 일자칸 값 직접 설정(JS) -> [조회]
        ③ [-][+] 버튼으로 이동 -> [조회]
        ④ 키보드 호환 모드(옵션 켜짐 시) -> [조회]
        최종적으로 화면 제목의 날짜가 목표 날짜와 같은지 검증합니다.
        """
        self.check_stop()
        frame = self.find_active_frame()
        current_date = self.get_screen_date()

        self.logger.log(f"날짜 이동 확인: 현재 화면={current_date}, 목표={target_date}")
        if current_date == target_date:
            self.logger.log(f"이미 목표 날짜({target_date})입니다.")
            return True

        target_parts = target_date.split(".")
        t_year = int(target_parts[0])
        t_month = int(target_parts[1])
        t_day = int(target_parts[2])

        # --------------------------------------------------------------
        # 방법 ① 달력 아이콘 클릭을 통한 이동
        # --------------------------------------------------------------
        self.logger.log_step(f"달력 아이콘을 이용한 날짜 이동 시도 ({target_date})")
        try:
            cal_clicked = frame.evaluate("""() => {
                const imgs = Array.from(document.querySelectorAll('img, button, i, span, a'));
                for (const el of imgs) {
                    const src = (el.src || el.className || el.getAttribute('aria-label') || '').toLowerCase();
                    const alt = (el.alt || '').toLowerCase();
                    if (src.includes('cal') || alt.includes('달력') || src.includes('date')) {
                        const r = el.getBoundingClientRect();
                        if (r.width > 10 && r.height > 10 && r.top < 150) {
                            el.click();
                            return true;
                        }
                    }
                }
                return false;
            }""")

            if cal_clicked:
                self.sleep_safe(0.4)
                # 달력 팝업 내 월 이동 및 날짜 선택
                frame.evaluate(f"""() => {{
                    // 달력 연/월 확인 및 화살표 조작
                    const targetYear = {t_year};
                    const targetMonth = {t_month};
                    const targetDay = {t_day};

                    // 날짜 셀들 탐색: 앞뒤 지난달/다음달 회색 날짜를 제외하고 첫 '1'부터 순서대로 이번 달 날짜 계산
                    const dayCells = Array.from(document.querySelectorAll('td, div[class*="day"], span[class*="day"]')).filter(c => {{
                        const r = c.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    }});

                    let foundFirst = false;
                    let currentDayCount = 1;
                    for (const c of dayCells) {{
                        const t = (c.innerText || '').trim();
                        if (t === '1') foundFirst = true;
                        if (foundFirst) {{
                            if (t === String(targetDay) && currentDayCount === targetDay) {{
                                c.click();
                                return true;
                            }}
                            if (/^\\d+$/.test(t)) currentDayCount++;
                        }}
                    }}
                    return false;
                }}""")
                self.sleep_safe(0.3)
                self.click_inquiry_button(frame)
                self.wait_table_loaded(frame)
                if self.get_screen_date() == target_date:
                    self.logger.log(f"달력 아이콘으로 날짜 이동 성공: {target_date}")
                    return True
        except Exception as e:
            self.logger.log(f"방법 ① (달력) 이동 실패: {e}", level="DEBUG")

        # --------------------------------------------------------------
        # 방법 ② 일자 입력칸 직접 JS 값 설정
        # --------------------------------------------------------------
        self.logger.log_step(f"자바스크립트 값 설정을 통한 날짜 이동 시도 ({target_date})")
        try:
            formatted_date_dot = f"{target_date}."
            js_set_result = frame.evaluate(f"""() => {{
                const inputs = Array.from(document.querySelectorAll('input[type="text"], input:not([type])'));
                for (const inp of inputs) {{
                    const val = (inp.value || '').trim();
                    if (/\\d{{4}}\\.\\d{{2}}\\.\\d{{2}}/.test(val) || inp.id.includes('date') || inp.name.includes('date')) {{
                        const proto = HTMLInputElement.prototype;
                        const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
                        if (setter) setter.call(inp, '{formatted_date_dot}');
                        else inp.value = '{formatted_date_dot}';
                        inp.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        inp.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        inp.dispatchEvent(new Event('blur', {{ bubbles: true }}));
                        return true;
                    }}
                }}
                return false;
            }}""")

            if js_set_result:
                self.sleep_safe(0.3)
                # 조회 버튼 최대 3회 시도
                for attempt in range(3):
                    self.click_inquiry_button(frame)
                    self.wait_table_loaded(frame)
                    if self.get_screen_date() == target_date:
                        self.logger.log(f"자바스크립트 설정으로 날짜 이동 성공: {target_date}")
                        return True
                    self.sleep_safe(0.5)
        except Exception as e:
            self.logger.log(f"방법 ② (JS 값 설정) 실패: {e}", level="DEBUG")

        # --------------------------------------------------------------
        # 방법 ③ [-][+] 버튼으로 이동
        # --------------------------------------------------------------
        self.logger.log_step(f"[-][+] 버튼을 통한 날짜 이동 시도 ({target_date})")
        try:
            # 일자칸 옆 [-], [+] 버튼 탐색
            step_btn_result = frame.evaluate(f"""() => {{
                const btns = Array.from(document.querySelectorAll('button, a, span, input[type="button"]'));
                let minusBtn = null, plusBtn = null;
                for (const b of btns) {{
                    const t = (b.innerText || b.value || '').trim();
                    const r = b.getBoundingClientRect();
                    if (r.top < 150 && r.width > 10 && r.height > 10) {{
                        if (t === '-' || t.includes('minus') || b.className.includes('prev')) minusBtn = b;
                        if (t === '+' || t.includes('plus') || b.className.includes('next')) plusBtn = b;
                    }}
                }}
                return {{ hasMinus: !!minusBtn, hasPlus: !!plusBtn }};
            }}""")

            if step_btn_result.get('hasMinus') or step_btn_result.get('hasPlus'):
                # 하루씩 이동하며 확인
                for _ in range(15):
                    self.check_stop()
                    curr = self.get_screen_date()
                    if curr == target_date:
                        self.click_inquiry_button(frame)
                        self.wait_table_loaded(frame)
                        if self.get_screen_date() == target_date:
                            self.logger.log(f"[-][+] 버튼으로 날짜 이동 성공: {target_date}")
                            return True
                        break
                    if not curr:
                        break
                    # 날짜 비교 후 클릭
                    if curr > target_date:
                        frame.evaluate("""() => {
                            const btns = Array.from(document.querySelectorAll('button, a, span, input[type="button"]'));
                            for (const b of btns) {
                                if ((b.innerText || '').trim() === '-' && b.getBoundingClientRect().top < 150) { b.click(); break; }
                            }
                        }""")
                    else:
                        frame.evaluate("""() => {
                            const btns = Array.from(document.querySelectorAll('button, a, span, input[type="button"]'));
                            for (const b of btns) {
                                if ((b.innerText || '').trim() === '+' && b.getBoundingClientRect().top < 150) { b.click(); break; }
                            }
                        }""")
                    self.sleep_safe(0.3)
        except Exception as e:
            self.logger.log(f"방법 ③ ([-][+] 버튼) 실패: {e}", level="DEBUG")

        # --------------------------------------------------------------
        # 방법 ④ 키보드 호환 모드 (옵션 켜짐 시 최후 수단)
        # --------------------------------------------------------------
        if self.options.get('keyboard_compat_mode', False):
            self.logger.log_step(f"키보드 호환 모드로 날짜 입력 시도 ({target_date})")
            try:
                frame.evaluate("""() => {
                    const inputs = Array.from(document.querySelectorAll('input[type="text"], input:not([type])'));
                    for (const inp of inputs) {
                        if (/\\d{4}\\.\\d{2}\\.\\d{2}/.test(inp.value || '') && inp.getBoundingClientRect().top < 150) {
                            inp.focus();
                            inp.select();
                            break;
                        }
                    }
                }""")
                self.page.keyboard.type(f"{target_date}.")
                self.sleep_safe(0.2)
                self.click_inquiry_button(frame)
                self.wait_table_loaded(frame)
                if self.get_screen_date() == target_date:
                    self.logger.log(f"키보드 호환 모드로 날짜 이동 성공: {target_date}")
                    return True
            except Exception as e:
                self.logger.log(f"방법 ④ (키보드 호환 모드) 실패: {e}", level="DEBUG")

        # 최종 검증
        final_date = self.get_screen_date()
        if final_date == target_date:
            return True

        self.logger.log(f"날짜 이동 최종 실패 (목표: {target_date}, 현재 화면: {final_date})", level="ERROR")
        return False

    def scroll_table(self, frame: Frame, direction: str):
        """명단 표를 위 또는 아래로 스크롤합니다."""
        frame.evaluate(f"""(dir) => {{
            const tables = Array.from(document.querySelectorAll('table, tbody, div[class*="grid"], div[class*="scroll"]'));
            for (const t of tables) {{
                if (t.scrollHeight > t.clientHeight && t.clientHeight > 100) {{
                    t.scrollTop += (dir === 'down' ? 180 : -180);
                    return;
                }}
            }}
            window.scrollBy(0, dir === 'down' ? 180 : -180);
        }}""", direction)
        self.sleep_safe(0.4)

    def find_and_ensure_visible_cell(self, frame: Frame, student_num: int, col_name: str) -> Dict[str, Any]:
        """
        학생의 특정 칸(마감 또는 교시)을 찾고 완전히 보일 때까지 스크롤하여 좌표를 반환합니다.
        """
        max_scroll_attempts = 15
        last_visible_nums = None

        for attempt in range(max_scroll_attempts):
            self.check_stop()
            info = frame.evaluate(FIND_STUDENT_CELL_JS, {
                'studentNum': student_num,
                'colName': col_name
            })

            status = info.get('status')
            if status == 'OK':
                if info.get('isFullyVisible'):
                    return info
                else:
                    # 행이 위/아래에 걸쳐 가려져 있는 경우 중앙으로 스크롤
                    self.logger.log(f"{student_num}번 행 일부 가려짐 -> 중앙 스크롤 정렬", level="DEBUG")
                    frame.evaluate(f"""(sNum) => {{
                        const cells = Array.from(document.querySelectorAll('td, div, span'));
                        for (const c of cells) {{
                            if ((c.innerText || '').trim() === String(sNum)) {{
                                const r = c.closest('tr') || c.parentElement;
                                if (r) r.scrollIntoView({{ block: 'center', inline: 'nearest' }});
                                break;
                            }}
                        }}
                    }}""", student_num)
                    self.sleep_safe(0.4)
                    continue

            elif status == 'NUM_NOT_FOUND':
                vis = info.get('visibleNums', [])
                if not vis:
                    self.scroll_table(frame, 'down')
                    continue

                if last_visible_nums == vis:
                    # 더 이상 번호가 바뀌지 않으면 명단 끝 도달
                    self.logger.log(f"명단 끝 도달 ({student_num}번 학생 없음, 현재 화면 번호: {vis})", level="WARN")
                    return {'status': 'NOT_FOUND_END_OF_TABLE'}

                last_visible_nums = vis
                min_vis = vis[0]
                max_vis = vis[-1]

                if student_num > max_vis:
                    self.scroll_table(frame, 'down')
                elif student_num < min_vis:
                    self.scroll_table(frame, 'up')
                else:
                    self.scroll_table(frame, 'down')

            else:
                self.logger.log(f"칸 탐색 상태 오류: {status}", level="WARN")
                return info

        return {'status': 'TIMEOUT'}

    def handle_confirm_dialogs(self, frame: Frame, max_wait: float = 3.0) -> bool:
        """
        저장하시겠습니까? 또는 저장되었습니다 같은 확인창 모달의 [확인] 버튼을 클릭합니다.
        """
        start = time.time()
        handled_any = False
        while time.time() - start < max_wait:
            self.check_stop()
            try:
                res = frame.evaluate(HANDLE_CONFIRM_MODAL_JS)
                if res.get('clicked'):
                    self.logger.log(f"확인창 처리: '{res.get('text')}' 버튼 클릭 완료")
                    handled_any = True
                    self.sleep_safe(0.3)
                    continue
            except Exception:
                pass
            time.sleep(0.2)
        return handled_any

    def handle_attendance_popup(self, frame: Frame, division: str, kind: str, default_reason: str, is_test_mode: bool) -> bool:
        """
        [마감]칸 클릭 시 열리는 '출결마감구분' 팝업에서 구분/종류/사유를 설정하고 적용/닫기를 클릭합니다.
        """
        self.check_stop()
        self.sleep_safe(0.5)

        # 팝업 열림 대기
        popup_opened = False
        for _ in range(10):
            has_title = frame.evaluate("""() => {
                const elems = Array.from(document.querySelectorAll('span, div, h1, h2, h3, h4, p, label'));
                return elems.some(el => (el.innerText || '').trim().includes('출결마감구분'));
            }""")
            if has_title:
                popup_opened = True
                break
            time.sleep(0.2)

        if not popup_opened:
            self.logger.log("출결마감구분 팝업이 열리지 않았습니다.", level="ERROR")
            return False

        self.logger.log(f"팝업 입력: 구분='{division}', 종류='{kind}', 사유='{default_reason}'")

        # 1. 구분 라디오 선택 (질병 / 미인정 / 기타 / 출석인정)
        div_selected = frame.evaluate(f"""(divName) => {{
            const labels = Array.from(document.querySelectorAll('label, span, td, div'));
            for (const l of labels) {{
                if ((l.innerText || '').trim() === divName) {{
                    const radio = l.querySelector('input[type="radio"]') || document.getElementById(l.htmlFor);
                    if (radio) {{
                        radio.click();
                        radio.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        return true;
                    }}
                    l.click();
                    return true;
                }}
            }}
            return false;
        }}""", division)

        if not div_selected:
            self.logger.log(f"구분 라디오('{division}') 선택 실패", level="WARN")

        # 구분을 고르면 종류 라디오가 활성화되므로 잠시 대기
        self.sleep_safe(0.4)

        # 2. 종류 라디오 선택 (지각 / 조퇴 / 결석 / 결과)
        kind_selected = frame.evaluate(f"""(kindName) => {{
            const labels = Array.from(document.querySelectorAll('label, span, td, div'));
            for (const l of labels) {{
                if ((l.innerText || '').trim() === kindName) {{
                    const radio = l.querySelector('input[type="radio"]') || document.getElementById(l.htmlFor);
                    if (radio) {{
                        radio.click();
                        radio.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        return true;
                    }}
                    l.click();
                    return true;
                }}
            }}
            return false;
        }}""", kind)

        if not kind_selected:
            self.logger.log(f"종류 라디오('{kind}') 선택 실패", level="WARN")

        # 3. 사유 입력 (GUI의 기본 사유 값 사용, 안전한 JS 설정)
        if default_reason:
            frame.evaluate(f"""(reasonText) => {{
                const inputs = Array.from(document.querySelectorAll('input[type="text"], textarea'));
                for (const inp of inputs) {{
                    // 팝업 내 사유 입력칸 탐색
                    const r = inp.getBoundingClientRect();
                    if (r.width > 50 && r.height > 15) {{
                        const proto = inp instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
                        const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
                        if (setter) setter.call(inp, reasonText);
                        else inp.value = reasonText;
                        inp.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        inp.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        break;
                    }}
                }}
            }}""", default_reason)

        self.sleep_safe(0.3)

        # 4. 버튼 클릭
        # 테스트 모드면 '적용' 대신 '닫기' 클릭
        # 운영 모드면 '적용' 클릭
        btn_to_click = "닫기" if is_test_mode else "적용"
        self.logger.log_step(f"팝업 [{btn_to_click}] 클릭 (테스트 모드={is_test_mode})")

        btn_clicked = frame.evaluate(f"""(btnText) => {{
            const buttons = Array.from(document.querySelectorAll('button, a, input[type="button"], [role="button"]'));
            for (const b of buttons) {{
                const t = (b.innerText || b.value || '').trim();
                if (t === btnText) {{
                    b.click();
                    return true;
                }}
            }}
            return false;
        }}""", btn_to_click)

        if not btn_clicked:
            self.logger.log(f"팝업 [{btn_to_click}] 버튼 클릭 실패", level="ERROR")
            return False

        # 팝업이 닫힐 때까지 대기하며 확인창 처리
        self.sleep_safe(0.5)
        self.handle_confirm_dialogs(frame)
        return True

    def click_period_cell(self, frame: Frame, student_num: int, period_name: str, double_click: bool = False) -> Tuple[bool, str]:
        """
        교시칸 클릭 규칙:
        1) 해당 학생의 [마감]칸에 값이 보일 때까지 대기
        2) 이미 '/' 표시면 누르지 않음
        3) 실제 마우스 좌표로 딱 1회 클릭 (또는 더블클릭)
        4) 마우스를 즉시 치우고 포커스 해제
        5) 최대 4초 동안 '/' 표시 또는 확인창 관찰
        6) 반응이 없어도 절대 재클릭 금지!
        """
        self.check_stop()
        self.logger.log_step(f"{student_num}번 {period_name}칸 1회 클릭 처리")

        # 1. 마감칸에 값이 나타날 때까지 대기 (명단 재렌더링 대기)
        magam_ready = False
        for _ in range(8):
            m_info = frame.evaluate(FIND_STUDENT_CELL_JS, {
                'studentNum': student_num,
                'colName': 'magam'
            })
            if m_info.get('status') == 'OK' and m_info.get('cellText'):
                magam_ready = True
                break
            time.sleep(0.3)

        if not magam_ready:
            self.logger.log(f"{student_num}번 마감칸 텍스트 반영 지연", level="DEBUG")

        # 2. 교시칸 위치 및 현재 상태 확인
        cell_info = self.find_and_ensure_visible_cell(frame, student_num, period_name)
        if cell_info.get('status') != 'OK':
            return False, f"교시칸 탐색 실패 ({cell_info.get('status')})"

        # 이미 '/'인지 확인
        initial_val = cell_info.get('cellText', '')
        if '/' in initial_val:
            self.logger.log(f"{student_num}번 {period_name}칸은 이미 '/' 상태이므로 누르지 않습니다.")
            return True, "이미체크됨"

        # 클릭 전 행 전체 교시 상태 기록
        vals_before = frame.evaluate(GET_ROW_PERIOD_VALUES_JS, student_num)
        self.logger.log(f"클릭 전 {student_num}번 행 상태: {vals_before}")

        # 3. 실제 마우스 좌표로 정확히 1회 클릭
        rect = cell_info['cellRect']
        click_x = rect['x'] + rect['width'] / 2
        click_y = rect['y'] + rect['height'] / 2

        click_count = 2 if double_click else 1
        self.page.mouse.click(click_x, click_y, click_count=click_count)
        self.logger.log(f"마우스 클릭 수행 (좌표: {int(click_x)}, {int(click_y)}, 횟수: {click_count})")

        # 4. 마우스 포인터를 즉시 치워 호버 스타일 착시 방지
        self.page.mouse.move(0, 0)
        frame.evaluate("() => { if (document.activeElement) document.activeElement.blur(); }")

        # 5. 최대 4초 동안 '/' 표시 또는 확인창 관찰 (재클릭 절대 금지!)
        click_success = False
        start_wait = time.time()
        while time.time() - start_wait < 4.0:
            self.check_stop()
            self.handle_confirm_dialogs(frame, max_wait=0.3)

            vals_check = frame.evaluate(GET_ROW_PERIOD_VALUES_JS, student_num)
            curr_target_val = vals_check.get(period_name, '')
            if '/' in curr_target_val:
                click_success = True
                break
            time.sleep(0.3)

        # 클릭 후 행 전체 교시 상태 기록 및 스크린샷 저장
        vals_after = frame.evaluate(GET_ROW_PERIOD_VALUES_JS, student_num)
        self.logger.log(f"클릭 후 {student_num}번 행 상태: {vals_after}")
        self.logger.save_screenshot(self.page, f"period_click_{student_num}_{period_name}")

        if click_success:
            self.logger.log(f"{student_num}번 {period_name}칸 정상 반영 ('/' 확인됨)")
            return True, "정상반영"
        else:
            self.logger.log(f"{student_num}번 {period_name}칸 반응 없음 (안전 규칙: 재클릭 금지)", level="WARN")
            return False, "반응없음(재클릭금지)"

    def process_record(self, frame: Frame, record: Dict[str, Any]) -> Tuple[str, str]:
        """
        한 학생의 묶인 출결 기록을 나이스에 입력합니다.
        반환: (결과상태, 메시지)
        """
        self.check_stop()
        student_num = record['student_num']
        division = record['division']
        kind = record['kind']
        periods = record['periods']
        target_period = record.get('target_period')
        is_test_mode = self.options.get('test_mode', True)
        default_reason = self.options.get('default_reason', '')
        double_click = self.options.get('period_double_click', False)

        self.logger.log(f"학생 처리 시작: {student_num}번 [{division}/{kind}] 교시: {','.join(periods)}")

        # 1. 학생 행 노출 확인 및 이름 읽기
        cell_info = self.find_and_ensure_visible_cell(frame, student_num, 'magam')
        if cell_info.get('status') != 'OK':
            err_msg = f"학생 행을 찾을 수 없음 ({cell_info.get('status')})"
            self.logger.log(f"{student_num}번 {err_msg}", level="ERROR")
            return "실패", err_msg

        student_name = cell_info.get('studentName', '')
        if student_name:
            self.logger.log(f"{student_num}번 화면 이름 확인: {student_name}")

        # 2. 결과(과목별) 출결인 경우: 교시 목록의 각 교시칸을 직접 클릭하여 팝업 처리
        if kind == '결과':
            self.logger.log(f"{student_num}번 결과(교시별) 처리: {periods}")
            for p in periods:
                self.check_stop()
                p_info = self.find_and_ensure_visible_cell(frame, student_num, p)
                if p_info.get('status') != 'OK':
                    return "실패", f"{p}칸 탐색 실패"

                p_rect = p_info['cellRect']
                self.page.mouse.click(p_rect['x'] + p_rect['width'] / 2, p_rect['y'] + p_rect['height'] / 2)
                self.sleep_safe(0.4)

                ok = self.handle_attendance_popup(frame, division, kind, default_reason, is_test_mode)
                if not ok:
                    return "실패", f"{p} 팝업 처리 실패"

            status = "테스트완료" if is_test_mode else "완료"
            return status, f"결과 교시별({','.join(periods)}) 처리 성공"

        # 3. 결석 / 지각 / 조퇴인 경우: [마감]칸 클릭하여 팝업 처리
        self.logger.log_step(f"{student_num}번 [마감]칸 클릭")
        rect = cell_info['cellRect']
        self.page.mouse.click(rect['x'] + rect['width'] / 2, rect['y'] + rect['height'] / 2)
        self.sleep_safe(0.4)

        popup_ok = self.handle_attendance_popup(frame, division, kind, default_reason, is_test_mode)
        if not popup_ok:
            # 1회 재시도 (팝업 닫고 다시 클릭)
            self.logger.log(f"{student_num}번 팝업 처리 실패 -> 1회 재시도", level="WARN")
            frame.evaluate("""() => {
                const closeBtns = Array.from(document.querySelectorAll('button, a'));
                for (const b of closeBtns) {
                    if ((b.innerText || '').trim() === '닫기') { b.click(); break; }
                }
            }""")
            self.sleep_safe(0.6)
            self.page.mouse.click(rect['x'] + rect['width'] / 2, rect['y'] + rect['height'] / 2)
            self.sleep_safe(0.4)
            popup_ok = self.handle_attendance_popup(frame, division, kind, default_reason, is_test_mode)
            if not popup_ok:
                return "실패", "출결마감구분 팝업 처리 실패"

        # 4. 교시칸 처리 (마감 후)
        # - 결석: 교시칸 클릭 없음
        if kind == '결석':
            self.logger.log(f"{student_num}번 결석 처리 완료 (교시칸 클릭 없음)")
            status = "테스트완료" if is_test_mode else "완료"
            return status, "결석 마감칸 입력 완료"

        # - 지각 / 조퇴: 교시칸 1회 클릭 처리
        if target_period:
            if is_test_mode:
                self.logger.log(f"테스트 모드: {student_num}번 {target_period}칸 클릭 건너뜀")
                return "테스트완료", f"테스트 완료 (클릭 대상: {target_period})"
            else:
                if self.options.get('click_period_cell', True):
                    clk_ok, clk_msg = self.click_period_cell(frame, student_num, target_period, double_click)
                    status = "완료" if clk_ok else "실패"
                    return status, f"{target_period}칸 처리: {clk_msg}"
                else:
                    self.logger.log("옵션에 의해 교시칸 클릭이 비활성화되었습니다.")
                    return "완료", "마감칸만 처리됨(옵션)"

        status = "테스트완료" if is_test_mode else "완료"
        return status, "입력 완료"

    def click_save_button(self, frame: Frame) -> bool:
        """한 날짜 처리 후 상단 [저장] 버튼을 클릭하고 확인창을 처리합니다."""
        self.logger.log_step("상단 [저장] 버튼 클릭")
        try:
            clicked = frame.evaluate("""() => {
                const buttons = Array.from(document.querySelectorAll('button, a, input[type="button"], [role="button"]'));
                for (const b of buttons) {
                    const t = (b.innerText || b.value || '').trim();
                    if (t === '저장') {
                        const r = b.getBoundingClientRect();
                        if (r.width > 20 && r.height > 15 && r.top < 150) {
                            b.click();
                            return true;
                        }
                    }
                }
                return false;
            }""")
            if clicked:
                self.sleep_safe(0.5)
                # "저장하시겠습니까?" 및 "저장되었습니다" 연속 처리
                self.handle_confirm_dialogs(frame, max_wait=4.0)
                return True
        except Exception as e:
            self.logger.log(f"저장 버튼 클릭 실패: {e}", level="WARN")
        return False

    def run_automation(self, excel_data: Dict[str, Any]):
        """전체 자동화 실행 루프"""
        records_by_date = excel_data['records_by_date']
        errors = excel_data.get('errors', [])

        self.logger.log("=" * 60)
        self.logger.log(f"나이스 출결 자동 입력 시작 ({APP_VERSION})")
        self.logger.log(f"옵션: {self.options}")
        self.logger.log("=" * 60)

        # 엑셀 오류 행 사전 기록
        for err in errors:
            self.logger.record_csv(
                err.get('date', ''),
                err.get('num', 0),
                '', '',
                [err.get('period', '')],
                [err.get('row', 0)],
                "엑셀오류",
                err.get('error', '엑셀 파싱 오류')
            )

        frame = self.find_active_frame()

        for date_str, student_records in records_by_date.items():
            self.check_stop()
            self.logger.log("-" * 50)
            self.logger.log(f"날짜 처리 시작: {date_str} (총 {len(student_records)}명)")

            # 날짜 이동 시도 및 화면 제목 날짜 확인
            nav_ok = self.navigate_to_date(date_str)
            if not nav_ok:
                err_msg = f"화면 날짜 이동 실패 (목표: {date_str}) -> 해당 날짜 학생 건너뜀"
                self.logger.log(err_msg, level="ERROR")
                for rec in student_records:
                    self.logger.record_csv(
                        rec['date'], rec['student_num'], rec['division'], rec['kind'],
                        rec['periods'], rec['source_rows'], "실패", err_msg
                    )
                continue

            # 날짜 일치 확인 재검증
            current_screen_date = self.get_screen_date()
            if current_screen_date != date_str:
                err_msg = f"안전 규칙 위반: 화면 날짜({current_screen_date})와 목표({date_str}) 불일치 -> 전원 건너뜀"
                self.logger.log(err_msg, level="ERROR")
                for rec in student_records:
                    self.logger.record_csv(
                        rec['date'], rec['student_num'], rec['division'], rec['kind'],
                        rec['periods'], rec['source_rows'], "실패", err_msg
                    )
                continue

            # 날짜 내 학생 순회
            for rec in student_records:
                self.check_stop()
                try:
                    res_status, res_msg = self.process_record(frame, rec)
                except Exception as e:
                    res_status = "실패"
                    res_msg = f"예외 발생: {str(e)}"
                    self.logger.log(f"{rec['student_num']}번 처리 중 오류: {e}\n{traceback.format_exc()}", level="ERROR")
                    self.logger.save_screenshot(self.page, f"error_{rec['student_num']}")

                self.logger.record_csv(
                    rec['date'],
                    rec['student_num'],
                    rec['division'],
                    rec['kind'],
                    rec['periods'],
                    rec['source_rows'],
                    res_status,
                    res_msg
                )

            # 한 날짜의 학생을 모두 처리한 후 옵션에 따라 [저장] 클릭
            if self.options.get('auto_save_per_date', False) and not self.options.get('test_mode', True):
                self.logger.log(f"{date_str} 처리 완료 -> 자동 저장 클릭")
                self.click_save_button(frame)

        self.logger.log("=" * 60)
        self.logger.log("모든 날짜 및 학생 기록 처리가 완료되었습니다.")
        self.logger.log("=" * 60)

    def close(self):
        try:
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass

# ----------------------------------------------------------------------
# Tkinter GUI 메인 화면 구성
# ----------------------------------------------------------------------
class NeisAttendanceGui:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME} - {APP_VERSION}")
        self.root.geometry("860x780")
        self.root.minsize(780, 680)

        # 상태 변수
        self.log_queue = queue.Queue()
        self.rpa_thread = None
        self.rpa_engine = None
        self.excel_data = None
        self.logger = None

        # 세션별 로그 폴더 초기화 (exe 옆 logs\날짜시각)
        base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.session_log_dir = os.path.join(base_dir, "logs", timestamp)
        self.logger = RpaLogger(self.session_log_dir)
        self.logger.set_ui_callback(self.append_log)

        self._build_ui()
        self._check_log_queue()

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use('clam')

        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --------------------------------------------------------------
        # 1단계: 브라우저 실행 및 연결 설정
        # --------------------------------------------------------------
        group1 = ttk.LabelFrame(main_frame, text=" 1단계: 브라우저 실행 및 로그인 ", padding=10)
        group1.pack(fill=tk.X, pady=(0, 10))

        # 나이스 주소
        r1 = ttk.Frame(group1)
        r1.pack(fill=tk.X, pady=2)
        ttk.Label(r1, text="나이스 주소:", width=12).pack(side=tk.LEFT)
        self.url_var = tk.StringVar(value=DEFAULT_NEIS_URL)
        ttk.Entry(r1, textvariable=self.url_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 브라우저 선택, 포트, 실행 버튼
        r2 = ttk.Frame(group1)
        r2.pack(fill=tk.X, pady=4)
        ttk.Label(r2, text="브라우저:", width=12).pack(side=tk.LEFT)
        self.browser_var = tk.StringVar(value="Chrome")
        b_combo = ttk.Combobox(r2, textvariable=self.browser_var, values=["Chrome", "Edge"], state="readonly", width=10)
        b_combo.pack(side=tk.LEFT, padx=(5, 15))

        ttk.Label(r2, text="원격 포트:").pack(side=tk.LEFT)
        self.port_var = tk.IntVar(value=DEFAULT_CDP_PORT)
        ttk.Entry(r2, textvariable=self.port_var, width=8).pack(side=tk.LEFT, padx=5)

        self.btn_launch_browser = ttk.Button(r2, text="1. 전용 브라우저 실행", command=self.on_launch_browser)
        self.btn_launch_browser.pack(side=tk.LEFT, padx=15)

        guide_lbl = ttk.Label(group1, text="※ 브라우저가 열리면 인증서 로그인 후 [일일 출결 관리] 화면으로 이동해 주세요.", foreground="#0066cc")
        guide_lbl.pack(anchor=tk.W, pady=(4, 0))

        # --------------------------------------------------------------
        # 2단계: 엑셀 파일 및 옵션 설정
        # --------------------------------------------------------------
        group2 = ttk.LabelFrame(main_frame, text=" 2단계: 엑셀 파일 선택 및 자동 입력 옵션 ", padding=10)
        group2.pack(fill=tk.X, pady=(0, 10))

        # 엑셀 파일 선택
        r3 = ttk.Frame(group2)
        r3.pack(fill=tk.X, pady=2)
        ttk.Label(r3, text="출결 엑셀:", width=12).pack(side=tk.LEFT)
        self.excel_path_var = tk.StringVar()
        ttk.Entry(r3, textvariable=self.excel_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(r3, text="파일 찾기...", command=self.on_browse_excel).pack(side=tk.LEFT)

        # 학년도 및 기본 사유
        r4 = ttk.Frame(group2)
        r4.pack(fill=tk.X, pady=4)
        ttk.Label(r4, text="학년도:", width=12).pack(side=tk.LEFT)
        current_year = datetime.datetime.now().year
        self.year_var = tk.StringVar(value=str(current_year))
        ttk.Entry(r4, textvariable=self.year_var, width=8).pack(side=tk.LEFT, padx=(5, 20))

        ttk.Label(r4, text="기본 사유(선택):").pack(side=tk.LEFT)
        self.reason_var = tk.StringVar()
        ttk.Entry(r4, textvariable=self.reason_var, width=28).pack(side=tk.LEFT, padx=5)
        ttk.Label(r4, text="(새 엑셀에 사유가 없을 때 팝업에 입력됨)", foreground="#666666").pack(side=tk.LEFT, padx=5)

        # 옵션 체크박스
        r5 = ttk.Frame(group2)
        r5.pack(fill=tk.X, pady=(6, 2))
        self.test_mode_var = tk.BooleanVar(value=True)
        cb_test = ttk.Checkbutton(r5, text="테스트 모드 (팝업 '닫기', 교시 클릭 안 함 - 기본 권장)", variable=self.test_mode_var)
        cb_test.pack(side=tk.LEFT, padx=(0, 15))

        self.auto_save_var = tk.BooleanVar(value=False)
        cb_save = ttk.Checkbutton(r5, text="날짜별 처리 후 [저장] 자동 클릭", variable=self.auto_save_var)
        cb_save.pack(side=tk.LEFT, padx=(0, 15))

        self.slow_mode_var = tk.BooleanVar(value=False)
        cb_slow = ttk.Checkbutton(r5, text="느리게 실행 (대기시간 1.6배)", variable=self.slow_mode_var)
        cb_slow.pack(side=tk.LEFT)

        r6 = ttk.Frame(group2)
        r6.pack(fill=tk.X, pady=(2, 2))
        self.click_period_var = tk.BooleanVar(value=True)
        cb_period = ttk.Checkbutton(r6, text="마감 후 교시칸 클릭", variable=self.click_period_var)
        cb_period.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(r6, text="교시칸 클릭 방식:").pack(side=tk.LEFT)
        self.click_method_var = tk.StringVar(value="single")
        ttk.Radiobutton(r6, text="1회 클릭", variable=self.click_method_var, value="single").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(r6, text="더블클릭", variable=self.click_method_var, value="double").pack(side=tk.LEFT, padx=2)

        r7 = ttk.Frame(group2)
        r7.pack(fill=tk.X, pady=(2, 2))
        self.keyboard_mode_var = tk.BooleanVar(value=False)
        cb_kb = ttk.Checkbutton(r7, text="키보드 호환 모드 사용 (주의: 보안 프로그램과 충돌 가능)", variable=self.keyboard_mode_var)
        cb_kb.pack(side=tk.LEFT)

        # --------------------------------------------------------------
        # 조작 버튼 줄
        # --------------------------------------------------------------
        btn_bar = ttk.Frame(main_frame)
        btn_bar.pack(fill=tk.X, pady=(0, 10))

        self.btn_inspect = ttk.Button(btn_bar, text="엑셀 검사 및 미리보기", command=self.on_inspect_excel)
        self.btn_inspect.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_start = ttk.Button(btn_bar, text="2. 자동 입력 시작", command=self.on_start_automation)
        self.btn_start.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_stop = ttk.Button(btn_bar, text="중지", command=self.on_stop_automation, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=(0, 10))

        # 로그 폴더 열기 버튼
        ttk.Button(btn_bar, text="로그 폴더 열기", command=self.on_open_log_folder).pack(side=tk.RIGHT)

        # --------------------------------------------------------------
        # 로그 출력 창
        # --------------------------------------------------------------
        log_group = ttk.LabelFrame(main_frame, text=" 진행 상황 및 로그 ", padding=5)
        log_group.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_group, wrap=tk.WORD, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_text.tag_config("INFO", foreground="#222222")
        self.log_text.tag_config("STEP", foreground="#0055aa", font=("Consolas", 9, "bold"))
        self.log_text.tag_config("WARN", foreground="#cc6600")
        self.log_text.tag_config("ERROR", foreground="#cc0000", font=("Consolas", 9, "bold"))
        self.log_text.tag_config("DEBUG", foreground="#777777")

        self.append_log(f"[{APP_NAME}] 프로그램이 준비되었습니다. ({APP_VERSION})")
        self.append_log(f"로그 폴더: {self.session_log_dir}")

    def append_log(self, msg: str, level: str = "INFO"):
        self.log_queue.put((msg, level))

    def _check_log_queue(self):
        try:
            while True:
                msg, level = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, f"{msg}\n", level)
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(100, self._check_log_queue)

    def on_launch_browser(self):
        """1단계: 원격 제어 브라우저 실행"""
        b_type = self.browser_var.get()
        port = self.port_var.get()
        url = self.url_var.get()

        try:
            launch_remote_browser(b_type, port, url)
            self.append_log(f"[{b_type}] 브라우저가 원격 제어 모드(포트 {port})로 실행되었습니다.", "STEP")
            self.append_log("브라우저에서 로그인 후 '일일 출결 관리' 화면으로 이동해 주세요.", "INFO")
        except Exception as e:
            messagebox.showerror("브라우저 실행 오류", str(e))
            self.append_log(f"브라우저 실행 실패: {e}", "ERROR")

    def on_browse_excel(self):
        filename = filedialog.askopenfilename(
            title="출결 엑셀 파일 선택",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
        )
        if filename:
            self.excel_path_var.set(filename)
            self.on_inspect_excel()

    def on_inspect_excel(self):
        """엑셀 검사 및 미리보기"""
        filepath = self.excel_path_var.get().strip()
        if not filepath:
            messagebox.showwarning("안내", "엑셀 파일을 먼저 선택해 주세요.")
            return

        try:
            self.excel_data = read_attendance_excel(filepath)
            self.append_log("-" * 60, "INFO")
            self.append_log(f"[엑셀 검사 결과] 파일: {os.path.basename(filepath)}", "STEP")
            self.append_log(f"- 머리글 행: {self.excel_data['header_row']}행")
            self.append_log(f"- 매핑된 열: {self.excel_data['col_map']}")
            self.append_log(f"- 총 데이터 행: {self.excel_data['total_rows']}개 행 -> {self.excel_data['total_records']}개 학생 기록으로 묶임")

            for d, records in self.excel_data['records_by_date'].items():
                self.append_log(f"▶ [{d}] ({len(records)}건):")
                for r in records:
                    periods_str = ",".join(r['periods'])
                    target_str = f"→ {r['target_period']}칸 클릭" if r['target_period'] else ("→ 각 교시칸 팝업 입력" if r['kind'] == '결과' else "→ 마감칸만 처리(결석)")
                    grade_str = f"({r['student_grade_class']}) " if r['student_grade_class'] else ""
                    self.append_log(f"   {d} {r['student_num']}번 {grade_str}[{r['division']}/{r['kind']}] 교시: {periods_str} {target_str}")

            if self.excel_data['errors']:
                self.append_log(f"※ 엑셀 인식 오류/누락 ({len(self.excel_data['errors'])}건):", "WARN")
                for err in self.excel_data['errors']:
                    self.append_log(f"   {err['row']}행: {err['error']} (일자={err['date']}, 번호={err['num']}, 교시={err['period']}, 출결={err['status']})", "WARN")

            self.append_log("-" * 60, "INFO")
            messagebox.showinfo("검사 완료", f"총 {self.excel_data['total_rows']}개 행을 {self.excel_data['total_records']}개 학생 기록으로 성공적으로 묶었습니다.")

        except Exception as e:
            self.append_log(f"엑셀 검사 실패: {e}\n{traceback.format_exc()}", "ERROR")
            messagebox.showerror("엑셀 검사 실패", str(e))

    def on_start_automation(self):
        """자동 입력 시작"""
        filepath = self.excel_path_var.get().strip()
        if not filepath:
            messagebox.showwarning("안내", "엑셀 파일을 선택해 주세요.")
            return

        if not self.excel_data:
            try:
                self.excel_data = read_attendance_excel(filepath)
            except Exception as e:
                messagebox.showerror("엑셀 오류", str(e))
                return

        total_records = self.excel_data['total_records']
        if total_records == 0:
            messagebox.showwarning("안내", "입력할 학생 기록이 없습니다.")
            return

        is_test = self.test_mode_var.get()
        mode_str = "테스트 모드 (실제 저장 및 교시 클릭 안 함)" if is_test else "★운영 모드 (실제 나이스에 입력 반영)★"

        confirm = messagebox.askyesno(
            "입력 시작 확인",
            f"총 {total_records}건의 학생 출결 기록을 처리합니다.\n\n"
            f"실행 모드: {mode_str}\n\n"
            f"브라우저에서 나이스 일일 출결 관리 화면이 열려 있는지 확인하셨습니까?\n"
            f"자동 입력을 시작하시겠습니까?"
        )
        if not confirm:
            return

        # UI 상태 변경
        self.btn_start.config(state=tk.DISABLED)
        self.btn_inspect.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)

        options = {
            'test_mode': is_test,
            'auto_save_per_date': self.auto_save_var.get(),
            'slow_mode': self.slow_mode_var.get(),
            'click_period_cell': self.click_period_var.get(),
            'period_double_click': (self.click_method_var.get() == "double"),
            'keyboard_compat_mode': self.keyboard_mode_var.get(),
            'default_reason': self.reason_var.get().strip(),
            'school_year': self.year_var.get().strip(),
        }

        self.rpa_engine = NeisAttendanceRPA(self.logger, options)

        # 백그라운드 스레드에서 자동화 실행
        def run_worker():
            try:
                port = self.port_var.get()
                ok = self.rpa_engine.connect_browser(port)
                if not ok:
                    self.append_log("브라우저 연결에 실패하여 작업을 중단합니다.", "ERROR")
                    return

                self.rpa_engine.run_automation(self.excel_data)
                self.append_log("자동화 작업이 성공적으로 종료되었습니다.", "STEP")
                messagebox.showinfo("작업 완료", "나이스 출결 자동 입력이 완료되었습니다.\n로그와 처리결과 CSV를 확인하세요.")
            except InterruptedError:
                self.append_log("작업이 사용자에 의해 중지되었습니다.", "WARN")
                messagebox.showwarning("작업 중지", "자동화 작업이 중지되었습니다.")
            except Exception as e:
                self.append_log(f"작업 실행 중 치명적 오류: {e}\n{traceback.format_exc()}", "ERROR")
                messagebox.showerror("오류 발생", f"실행 중 오류가 발생했습니다: {e}")
            finally:
                if self.rpa_engine:
                    self.rpa_engine.close()
                self.root.after(0, self._on_automation_finished)

        self.rpa_thread = threading.Thread(target=run_worker, daemon=True)
        self.rpa_thread.start()

    def _on_automation_finished(self):
        self.btn_start.config(state=tk.NORMAL)
        self.btn_inspect.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)

    def on_stop_automation(self):
        if self.rpa_engine:
            self.rpa_engine.request_stop()
            self.append_log("중지 명령을 전송했습니다. 현재 단계 완료 후 즉시 멈춥니다...", "WARN")

    def on_open_log_folder(self):
        if os.path.exists(self.session_log_dir):
            os.startfile(self.session_log_dir)
        else:
            messagebox.showinfo("안내", "로그 폴더가 아직 생성되지 않았습니다.")

# ----------------------------------------------------------------------
# 프로그램 진입점
# ----------------------------------------------------------------------
def main():
    root = tk.Tk()
    app = NeisAttendanceGui(root)
    root.mainloop()

if __name__ == '__main__':
    main()
