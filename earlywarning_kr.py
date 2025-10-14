'''최종: region_data 파일을 불러와서 지역별 지표별 상황을 표로 표시하고, 해당되는 이미지를 2*9 배열로 보여줌'''
import os
import re
import pandas as pd
import streamlit as st
from PIL import Image

st.set_page_config(page_title="일자리상황판 - 지역/구분 뷰어", layout="wide")
st.title("전국-일자리 조기경보서비스")

EXCEL_PATH = "region_data.xlsx"

if not os.path.exists(EXCEL_PATH):
    st.warning("엑셀 파일이 보이지 않아요. 위 경로에 'region_data.xlsx'가 있는지 확인해주세요.")
    st.stop()

# --------- 유틸 함수 ---------
def _normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _find_col(candidates, columns):
    for cand in candidates:
        for col in columns:
            if col.lower() == cand.lower():
                return col
    for col in columns:
        lc = col.lower()
        if any(cand.lower() in lc for cand in candidates):
            return col
    return None

def _drop_serial_cols(df):
    """연번/번호/Unnamed 등 불필요한 연번 컬럼 제거"""
    serial_patterns = [r'^\s*연번\s*$', r'^\s*번호\s*$', r'^\s*no\.?\s*$', r'^unnamed']
    cols = []
    for c in df.columns:
        lc = str(c).strip().lower()
        if any(re.search(p, lc) for p in serial_patterns):
            continue
        cols.append(c)
    return df[cols]

def _to_numeric_if_possible(s):
    try:
        if isinstance(s, str):
            return float(s.replace(",", ""))
        return float(s)
    except Exception:
        return s

# --------- 엑셀 로드 ---------
try:
    xl = pd.ExcelFile(EXCEL_PATH)
except Exception as e:
    st.error(f"엑셀 파일을 열 수 없습니다: {e}")
    st.stop()

def _pick_sheet(candidates):
    for name in xl.sheet_names:
        for c in candidates:
            if name.strip().lower() == c.lower():
                return name
    for name in xl.sheet_names:
        ln = name.strip().lower()
        if any(c.lower() in ln for c in candidates):
            return name
    return None

data_sheet  = _pick_sheet(["data", "데이터", "자료", "값"])
image_sheet = _pick_sheet(["image", "이미지", "img", "그림"])

if data_sheet is None:
    st.error("엑셀에 'data' 시트(또는 유사명)가 필요합니다.")
    st.stop()

try:
    df_data = pd.read_excel(EXCEL_PATH, sheet_name=data_sheet, header=0)
except Exception as e:
    st.error(f"data 시트를 읽는 중 오류: {e}")
    st.stop()

df_data = _normalize_columns(df_data)
df_data = _drop_serial_cols(df_data)

# '지역' 컬럼 추정
region_col = _find_col(["지역", "시도", "region"], df_data.columns)
if region_col is None:
    region_col = df_data.columns[0]

# 구분(카테고리) 컬럼들
category_cols = [c for c in df_data.columns if c != region_col]
if not category_cols:
    st.error("구분(카테고리)로 사용할 컬럼이 없습니다. data 시트의 1행 헤더를 확인해주세요.")
    st.stop()

# 이미지 매핑 로드(선택)
img_map = None
if image_sheet is not None:
    try:
        df_img = pd.read_excel(EXCEL_PATH, sheet_name=image_sheet, header=0)
        df_img = _normalize_columns(df_img)
        img_region_col = _find_col(["지역", "시도", "region"],  df_img.columns)
        img_cat_col    = _find_col(["구분", "항목", "category"], df_img.columns)
        img_file_col   = _find_col(["image", "img", "filename", "file", "파일", "파일명", "이미지", "경로", "path"], df_img.columns)
        if all([img_region_col, img_cat_col, img_file_col]):
            # 경로 정규화(역슬래시 -> 슬래시)까지 포함
            df_img[img_file_col] = df_img[img_file_col].astype(str).str.replace("\\", "/", regex=False)
            img_map = df_img[[img_region_col, img_cat_col, img_file_col]].dropna()
        else:
            st.warning("image 시트에서 (지역, 구분, 이미지파일명) 컬럼을 찾지 못했습니다. 컬럼명을 확인해주세요.")
    except Exception as e:
        st.warning(f"image 시트를 읽는 중 문제가 발생했습니다: {e}")

# --------- UI: 지역 선택 ---------
regions = df_data[region_col].dropna().astype(str).unique().tolist()

# --------- UI: 지역 선택 및 표 표시 (세로 배치) ---------
st.markdown("### 〓지역 선택〓")  
region = st.radio(
    "17개 시도",
    regions,
    index=0 if "경기도" not in regions else regions.index("경기도"),
    horizontal=True
)

# 선택 지역 행
row = df_data[df_data[region_col].astype(str) == str(region)]
if row.empty:
    st.error("선택한 지역에 해당하는 데이터가 없습니다.")
    st.stop()

# --------- DATA 표: 카테고리를 열로, 지역은 한 행으로 ---------
row_vals = {c: _to_numeric_if_possible(row.iloc[0][c]) for c in category_cols}
one_row_df = pd.DataFrame([row_vals])

def highlight_status(val):
    """정상·관심·위기 상태에 따라 색상 및 스타일 지정"""
    val_str = str(val).strip()
    if val_str == "정상":
        color = "black"
        weight = "bold"
    elif val_str == "관심":
        color = "#DAA520"  # 짙은 노란색
        weight = "bold"
    elif val_str == "주의":
        color = "red"
        weight = "bold"
    else:
        color = "black"
        weight = "normal"
    return f"color: {color}; font-weight: {weight};"

# --------- DATA 표 표시 ---------
st.markdown("### 〓지역별 일자리지표〓")
styled_df = one_row_df.style.map(highlight_status)
st.dataframe(styled_df, use_container_width=True, hide_index=True)

st.markdown("---")
st.subheader(f"'{region}' 지표별 상세보기")

# --------- 선택 지역의 모든 이미지 2×9 그리드로 표시 ---------
def _gather_region_images(region_value):
    """선택 지역의 (구분, 이미지경로) 목록을 category_cols 순서대로 정리"""
    if img_map is None:
        return []

    def _norm(x):
        return str(x).strip().lower()

    # 지역 일치 (대소문자/공백 무시)
    reg_series = img_map.iloc[:, 0].astype(str)
    cat_series = img_map.iloc[:, 1].astype(str)
    file_series = img_map.iloc[:, 2].astype(str)

    mask_region = reg_series.str.strip().str.lower() == _norm(region_value)
    df_r = img_map[mask_region]
    if df_r.empty:
        return []

    # 카테고리별 대표 1개만 매핑(dict)
    mapping = {}
    for cat, path in zip(df_r.iloc[:,1].astype(str), df_r.iloc[:,2].astype(str)):
        if cat not in mapping:
            mapping[cat] = path

    # category_cols 순서대로 우선 정렬, 남는 항목은 뒤에
    ordered = []
    seen = set()
    for c in category_cols:
        # 정확 일치 우선, 없으면 느슨 매칭(포함)
        exact = [k for k in mapping.keys() if k == c]
        if exact:
            ordered.append((c, mapping[exact[0]]))
            seen.add(exact[0])
        else:
            partial = [k for k in mapping.keys() if _norm(c) in _norm(k) and k not in seen]
            if partial:
                ordered.append((c, mapping[partial[0]]))
                seen.add(partial[0])

    # 남은 것들 추가
    for k, v in mapping.items():
        if k not in seen:
            ordered.append((k, v))

    return ordered

entries = _gather_region_images(region)

if not entries:
    st.info("이 지역에 매핑된 이미지가 없습니다. image 시트를 확인해주세요.")
else:
    base_dir = os.path.dirname(os.path.abspath(EXCEL_PATH))
    # 최대 18개(2×9)만 출력
    max_items = min(18, len(entries))
    for i in range(0, max_items, 2):
        cols = st.columns(2)
        for j in range(2):
            idx = i + j
            if idx >= max_items:
                break
            cat, rel_path = entries[idx]
            # 경로 정규화 및 절대경로화
            rel_path = str(rel_path).replace("\\", "/")
            abs_path = rel_path if os.path.isabs(rel_path) else os.path.join(base_dir, rel_path)
            with cols[j]:
                if os.path.exists(abs_path):
                    st.image(Image.open(abs_path), use_container_width=True, caption=str(cat))
                else:
                    st.warning(f"이미지 없음: {abs_path}")


