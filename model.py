"""2023 한국은행 산업연관표: 브라우저(Pyodide)에서 실행하는 분석 엔진.

원본 Excel 파일의 계산된 셀 값(data_only)을 이용한다.
산업 간 거래표 탐색과 국산품 최종수요 변화의 3대 유발효과를 계산한다.
"""
import io
import json

import numpy as np
import openpyxl

book = {}
codes = []
names = []
sheets = []
L = None
M = None
V = None
residuals = {}


def _emit(obj):
    return json.dumps(obj, ensure_ascii=False, allow_nan=False)


def _code(value):
    """원본의 앞자리 0 보존: Excel에서는 코드가 문자열로 기록되어 있다."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _matrix(sheet_name):
    rows = book[sheet_name]
    by_row = {_code(row[0]): row for row in rows[6:] if len(row) > 1 and row[0] is not None}
    by_col = {_code(value): j for j, value in enumerate(rows[4]) if value is not None}
    try:
        m = np.array([[by_row[r][by_col[c]] for c in codes] for r in codes], dtype=float)
    except (KeyError, IndexError, TypeError, ValueError) as exc:
        raise ValueError(f"{sheet_name}: 산업별 행·열의 코드가 맞지 않습니다 ({exc}).") from exc
    if m.shape != (len(codes), len(codes)) or not np.isfinite(m).all():
        raise ValueError(f"{sheet_name}: 계수행렬에 결측치 또는 비정상 값이 있습니다.")
    return m


def initialize(raw, expected_count):
    """Excel을 로드하며 분류별로 산업 코드를 읽고 행렬 관계를 검증한다."""
    global book, codes, names, sheets, L, M, V, residuals
    required = (
        '총투입계수(A)', '국산투입계수(Ad)', '수입투입계수(Am)',
        '생산유발계수', '수입유발계수', '부가가치유발계수',
    )
    workbook = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    try:
        new_book = {name: list(workbook[name].values) for name in workbook.sheetnames}
    finally:
        workbook.close()
    missing = [name for name in required if name not in new_book]
    if missing:
        raise ValueError('필수 Excel 시트가 없습니다: ' + ', '.join(missing))
    book = new_book
    sheets = list(book)
    n = int(expected_count)
    base = book['국산투입계수(Ad)']
    codes = [_code(value) for value in base[4][2:2+n]]
    sector_names = {_code(row[0]): str(row[1]).strip() for row in base[6:]
                    if len(row) > 1 and row[0] is not None and row[1] is not None}
    if len(codes) != n or len(set(codes)) != n or not all(c in sector_names for c in codes):
        raise ValueError(f'산업코드가 예상한 {n}개 부문과 일치하지 않습니다.')
    names = [sector_names[c] for c in codes]
    L = _matrix('생산유발계수')
    M = _matrix('수입유발계수')
    V = _matrix('부가가치유발계수')
    A_d = _matrix('국산투입계수(Ad)')
    A_m = _matrix('수입투입계수(Am)')
    total = book['총투입계수(A)']
    value_added = next((row for row in total if row and _code(row[0]) == '9690'), None)
    if value_added is None:
        raise ValueError('부가가치계수 행(9690)을 찾지 못했습니다.')
    cindex = {_code(c): i for i, c in enumerate(total[4]) if c is not None}
    v = np.array([value_added[cindex[c]] for c in codes], dtype=float)
    identity = np.eye(n)
    residuals = {
        '생산 (I − Ad)L = I': float(np.max(np.abs((identity-A_d) @ L-identity))),
        '수입 AmL = M': float(np.max(np.abs(A_m @ L-M))),
        '부가가치 diag(v)L = V': float(np.max(np.abs(v[:, None] * L-V))),
    }
    if not all(np.isfinite(x) and x < 1e-7 for x in residuals.values()):
        raise ValueError('입출력 계수 검증에 실패했습니다: '+str(residuals))
    return _emit({'sectors': [{'code': c, 'name': name} for c, name in zip(codes, names)],
                  'sheets': sheets, 'validation': residuals, 'count': n})


def explore(sheet_name, rowcode=""):
    if sheet_name not in book:
        raise ValueError('해당 Excel 시트가 존재하지 않습니다.')
    rows = book[sheet_name]
    opts = [{'code': _code(row[0]), 'name': str(row[1]).strip()}
            for row in rows[6:] if len(row)>1 and row[0] is not None and row[1] is not None]
    if not opts:
        raise ValueError('조회 가능한 행이 없습니다.')
    selected = next((row for row in rows[6:] if row and _code(row[0]) == rowcode), None)
    if selected is None:
        code = opts[0]['code']
        selected = next(row for row in rows[6:] if row and _code(row[0]) == code)
    columns = []
    header = rows[4]
    titles = rows[5]
    for i in range(2, len(header)):
        if header[i] is None:
            continue
        value = selected[i] if i < len(selected) else None
        if isinstance(value, (int, float)):
            value = float(value)
            if not np.isfinite(value):
                value = None
        elif value is not None:
            value = str(value)
        columns.append({'code': _code(header[i]),
                        'name': str(titles[i] or '') if i < len(titles) else '',
                        'value': value})
    monetary = sheet_name.startswith('A표_') or sheet_name.startswith('최종수요')
    return _emit({'rows': opts, 'selected': _code(selected[0]), 'name': str(selected[1]),
                  'unit': '백만원' if monetary else '계수 (무차원)', 'cells': columns})


def simulate(payload_json):
    if L is None:
        raise ValueError('Excel을 먼저 불러오세요.')
    selected = json.loads(payload_json)
    f = np.zeros(len(codes))
    lookup = {c: i for i, c in enumerate(codes)}
    for item in selected:
        code = str(item['code'])
        amount = float(item['amount'])  # 백만원 단위
        if code not in lookup or not np.isfinite(amount) or amount < 0:
            raise ValueError('유효한 산업과 0 이상의 금액을 입력해야 합니다.')
        f[lookup[code]] += amount
    if f.sum() <= 0:
        raise ValueError('하나 이상의 양(+)의 최종수요 증가액을 입력하세요.')
    x = L @ f
    m = M @ f
    v = V @ f
    output = [{'code': code, 'name': name, 'production': float(px),
               'imports': float(im), 'valueAdded': float(va), 'finalDemand': float(fd)}
              for code, name, px, im, va, fd in zip(codes, names, x, m, v, f)]
    output.sort(key=lambda row: row['production'], reverse=True)
    return _emit({'demand': float(f.sum()), 'production': float(x.sum()),
                  'imports': float(m.sum()), 'valueAdded': float(v.sum()),
                  'productionMultiplier': float(x.sum()/f.sum()),
                  'rows': output})
