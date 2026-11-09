# -*- coding: utf-8 -*-
"""
test_excel_parser.py
neis_attendance_rpa.py 의 엑셀 읽기, 출결구분 해석, 학생별 기록 묶기 및 정렬 검증 테스트
"""
import os
import openpyxl
from neis_attendance_rpa import (
    read_attendance_excel,
    parse_status,
    parse_student_num,
    normalize_date,
    normalize_period,
    PERIOD_ORDER
)

def run_test():
    """예시 데이터를 엑셀로 작성하여 파싱, 묶기, 해석 결과 검증"""
    test_excel_path = "test_sample_attendance.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "출결기록"

    # 머리글 작성
    headers = ["일자", "학번", "교시", "출결구분", "상태분류", "시스템등록일시"]
    ws.append(headers)

    # 예시 데이터 행 추가
    rows = [
        ["2026-11-05", 17, "5교시", "인정결", "결석류", "2026-11-05 14:00"],
        ["2026-11-05", 17, "6교시", "인정결", "결석류", "2026-11-05 14:00"],
        ["2026-11-05", 17, "7교시", "인정결", "결석류", "2026-11-05 14:00"],
        ["2026-11-09", 9, "조례", "병지각", "지각류", "2026-11-09 09:00"],
        ["2026-11-09", 9, "1교시", "병지각", "지각류", "2026-11-09 09:00"],
        ["2026-11-09", 13, "3교시", "인정조퇴", "조퇴류", "2026-11-09 11:30"],
        ["2026-11-09", 13, "4교시", "인정조퇴", "조퇴류", "2026-11-09 11:30"],
        ["2026-11-09", 13, "5교시", "인정조퇴", "조퇴류", "2026-11-09 11:30"],
        ["2026-11-09", 13, "6교시", "인정조퇴", "조퇴류", "2026-11-09 11:30"],
    ]

    for row in rows:
        ws.append(row)

    wb.save(test_excel_path)
    wb.close()
    print(f"[OK] 테스트 엑셀 파일 생성: {test_excel_path}")

    # 엑셀 파싱 실행
    result = read_attendance_excel(test_excel_path)
    print("\n[엑셀 파싱 결과 요약]")
    print(f"- 머리글 행 번호: {result['header_row']}")
    print(f"- 매핑된 열 인덱스: {result['col_map']}")
    print(f"- 총 원본 행 수: {result['total_rows']}")
    print(f"- 묶인 학생 기록 수: {result['total_records']}")

    records_by_date = result['records_by_date']
    all_recs = []
    for d, recs in records_by_date.items():
        for r in recs:
            all_recs.append(r)
            print(f"날짜: {r['date']}, 학번: {r['student_num']}, 구분: {r['division']}, 종류: {r['kind']}, "
                  f"교시목록: {r['periods']}, 클릭대상: {r['target_period']}, 원본행: {r['source_rows']}")

    # 1. 17번 11/05: 출석인정 / 결석 / [5교시, 6교시, 7교시]
    rec_17 = [r for r in all_recs if r['student_num'] == 17 and '11.05' in r['date']][0]
    assert rec_17['division'] == '출석인정', f"17번 구분 오류: {rec_17['division']}"
    assert rec_17['kind'] == '결석', f"17번 종류 오류: {rec_17['kind']}"
    assert rec_17['periods'] == ['5교시', '6교시', '7교시'], f"17번 교시 목록 오류: {rec_17['periods']}"
    assert rec_17['target_period'] is None, f"17번 클릭대상 오류: {rec_17['target_period']}"

    # 2. 9번 11/09: 질병 / 지각 / [조회, 1교시] → 클릭 대상: 1교시 (마지막 교시)
    rec_9 = [r for r in all_recs if r['student_num'] == 9 and '11.09' in r['date']][0]
    assert rec_9['division'] == '질병', f"9번 구분 오류: {rec_9['division']}"
    assert rec_9['kind'] == '지각', f"9번 종류 오류: {rec_9['kind']}"
    assert rec_9['periods'] == ['조회', '1교시'], f"9번 교시 목록 오류: {rec_9['periods']}"
    assert rec_9['target_period'] == '1교시', f"9번 클릭대상 오류: {rec_9['target_period']}"

    # 3. 13번 11/09: 출석인정 / 조퇴 / [3교시, 4교시, 5교시, 6교시] → 클릭 대상: 3교시 (첫 교시)
    rec_13 = [r for r in all_recs if r['student_num'] == 13 and '11.09' in r['date']][0]
    assert rec_13['division'] == '출석인정', f"13번 구분 오류: {rec_13['division']}"
    assert rec_13['kind'] == '조퇴', f"13번 종류 오류: {rec_13['kind']}"
    assert rec_13['periods'] == ['3교시', '4교시', '5교시', '6교시'], f"13번 교시 목록 오류: {rec_13['periods']}"
    assert rec_13['target_period'] == '3교시', f"13번 클릭대상 오류: {rec_13['target_period']}"

    # 4. 학번 3자리 이상 (20917) 테스트
    num, grade_class = parse_student_num(20917)
    assert num == 17, f"학번 번호 파싱 오류: {num}"
    assert grade_class == "209", f"학번 학년반 파싱 오류: {grade_class}"

    # 5. 상태분류 불일치 에러 테스트
    div, kind, err = parse_status("병지각", "결석류")
    assert err is not None, "상태분류 불일치 감지 실패"

    # 6. 미인정 vs 인정 우선순위 테스트
    div_unauth, kind_unauth, err_unauth = parse_status("미인정결", "결석류")
    assert div_unauth == "미인정", f"미인정 판별 실패: {div_unauth}"
    assert kind_unauth == "결석", f"미인정결 종류 판별 실패: {kind_unauth}"

    print("\n[SUCCESS] 모든 단위 테스트 및 요구사항 검증이 성공적으로 완료되었습니다!")

if __name__ == '__main__':
    run_test()
