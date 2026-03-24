import re
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# ✅ 비밀번호 엑셀(스마트스토어) 복호화
import msoffcrypto

st.set_page_config(page_title="주문파일 → 송장파일 변환", layout="centered")
st.title("📦 주문파일 → 송장 출력용 파일 변환기 (자동 플랫폼 판별 + 다중 업로드 통합)")

st.markdown("""
- 쿠팡/스마트스토어/지마켓(Gmarket)/thirtymall(떠리몰) 주문 엑셀(xlsx) **여러 개를 한번에 업로드**
- 파일별로 **플랫폼 자동 판별**
- **헤더(컬럼명) 기반 자동 매핑**
- 결과는 **한 개의 송장파일로 통합 변환**
- ✅ 스마트스토어 파일: **항상 비밀번호 1234로 열기 + 첫 번째 행 제거 후 컬럼매칭**
- ✅ 스마트스토어 품목명: **Q열(상품명) + S열(옵션정보)**
- ✅ 쿠팡 품목명: **M열 노출상품명(옵션명)**
- ✅ 지마켓 품목명: **상품명 + 옵션/추가구성/사은품/덤**
- ✅ 지마켓 집하예정일: **발송마감일(없으면 발송예정일)**
- ✅ 지마켓 품목코드: **상품번호(없으면 판매자 관리코드)**
- ✅ thirtymall(떠리몰) 품목명: **S열(상품명) + V열(옵션명:옵션값)** (중복 글 1회 표기)
""")

# =========================
# 기본 송장 템플릿(첨부 송장파일.xlsx 기준 컬럼/순서 내장)
# =========================
DEFAULT_TEMPLATE_COLUMNS = [
    "고객주문번호",
    "집하예정일",
    "품목코드",
    "품목명",
    "기타1",
    "기타2",
    "내품수량",
    "박스수량",
    "받는분성명",
    "받는분전화번호",
    "받는분우편번호",
    "받는분주소(전체,분할)",
    "배송메세지1",
    "운송장번호",
]

def build_default_template_df() -> pd.DataFrame:
    return pd.DataFrame(columns=DEFAULT_TEMPLATE_COLUMNS)

# -------------------------
# 유틸: 컬럼명 정규화/검색
# -------------------------
def norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[()\-_/.,·:]", "", s)
    return s

def find_col(df: pd.DataFrame, candidates: list[str]):
    """df에서 candidates(후보 헤더명) 중 하나라도 일치/포함되면 해당 컬럼명 반환"""
    norm_cols = {norm(c): c for c in df.columns}

    # 1) 완전 일치
    for cand in candidates:
        nc = norm(cand)
        if nc in norm_cols:
            return norm_cols[nc]

    # 2) 부분 포함
    for df_norm, original in norm_cols.items():
        for cand in candidates:
            nc = norm(cand)
            if nc and (nc in df_norm or df_norm in nc):
                return original

    return None

def clean_series(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .fillna("")
        .replace(["nan", "None"], "", regex=False)
        .str.strip()
    )

# =========================
# ✅ 스마트스토어 암호(1234) 복호화 + 1행 제거(=header=1)
# =========================
SMARTSTORE_PASSWORD = "1234"

def decrypt_xlsx_if_needed(file_bytes: bytes, password: str) -> BytesIO:
    """
    암호화된 xlsx면 복호화해서 BytesIO 반환.
    암호화가 아니면 원본 BytesIO 반환.
    """
    bio = BytesIO(file_bytes)
    try:
        office = msoffcrypto.OfficeFile(bio)
        office.load_key(password=password)
        decrypted = BytesIO()
        office.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted
    except Exception:
        bio.seek(0)
        return bio

def read_excel_safely(uploaded_file, platform_hint: str | None = None) -> pd.DataFrame:
    """
    - smartstore: 비번 1234 복호화 + 첫 번째 행 제거 후(header=1) 로드
    - others: 일반 로드
    """
    file_bytes = uploaded_file.getvalue()

    if platform_hint == "smartstore":
        decrypted = decrypt_xlsx_if_needed(file_bytes, SMARTSTORE_PASSWORD)
        # ✅ 첫 번째 행 삭제 후 컬럼 매칭(2번째 행을 헤더로)
        return pd.read_excel(decrypted, header=1, dtype=str)

    return pd.read_excel(BytesIO(file_bytes), dtype=str)

# -------------------------
# 플랫폼 판별
# -------------------------
PLATFORM_SIGNATURES = {
    "coupang": [
        "노출상품명", "노출상품명(옵션명)", "등록상품명", "수취인이름", "주문번호", "결제액", "구매수"
    ],
    "smartstore": [
        "상품주문번호", "수취인명", "배송메시지", "배송메세지", "옵션정보", "우편번호"
    ],
    "gmarket": [
        "판매아이디", "주문번호", "상품명", "수령인명", "수령인 휴대폰", "우편번호",
        "주소", "배송시 요구사항", "발송마감일", "택배사명(발송방법)"
    ],
    # ✅ thirtymall(떠리몰) - 첨부파일 기준 컬럼 포함
    "thirtymall": [
        "쇼핑몰구분", "주문구분", "주문메모", "업무메시지",
        "상품명", "옵션명:옵션값", "수령자명", "수령자연락처", "우편번호", "주소"
    ],
}

def detect_platform(df: pd.DataFrame) -> str:
    cols_norm = set(norm(c) for c in df.columns)

    def score(keys):
        s = 0
        for k in keys:
            nk = norm(k)
            if nk in cols_norm:
                s += 2
            else:
                for c in cols_norm:
                    if nk and (nk in c or c in nk):
                        s += 1
                        break
        return s

    coupang_score = score(PLATFORM_SIGNATURES["coupang"])
    smart_score = score(PLATFORM_SIGNATURES["smartstore"])
    gmarket_score = score(PLATFORM_SIGNATURES["gmarket"])
    thirty_score = score(PLATFORM_SIGNATURES["thirtymall"])

    # ✅ 지마켓 파일은 판매아이디/제휴사명 값에 "지마켓/Gmarket/옥션"이 들어오는 케이스가 많음
    seller_col = find_col(df, ["판매아이디", "제휴사명", "판매채널", "쇼핑몰구분", "몰구분"])
    if seller_col is not None:
        vals = clean_series(df[seller_col]).str.lower()
        if (
            vals.str.contains("지마켓", na=False)
            | vals.str.contains("gmarket", na=False)
            | vals.str.contains("옥션", na=False)
            | vals.str.contains("auction", na=False)
        ).any():
            gmarket_score += 3

    # ✅ 떠리몰 파일은 '쇼핑몰구분' 값에 "떠리몰"/"thirtymall"이 들어오는 케이스가 많아서 값 기반 보정
    mall_col = find_col(df, ["쇼핑몰구분", "쇼핑몰", "mall", "shop"])
    if mall_col is not None:
        vals = clean_series(df[mall_col]).str.lower()
        if (vals.str.contains("떠리몰", na=False) | vals.str.contains("thirtymall", na=False)).any():
            thirty_score += 3

    if coupang_score == 0 and smart_score == 0 and gmarket_score == 0 and thirty_score == 0:
        return "unknown"

    scores = {
        "coupang": coupang_score,
        "smartstore": smart_score,
        "gmarket": gmarket_score,
        "thirtymall": thirty_score,
    }
    return max(scores, key=scores.get)

# -------------------------
# 자동 매핑 후보(템플릿 컬럼명 기준)
# -------------------------
CANDIDATES = {
    "고객주문번호": {
        "coupang": ["주문번호", "고객주문번호", "order number", "orderno"],
        "smartstore": ["상품주문번호", "상품 주문번호", "주문번호", "주문관리번호", "order no"],
        "gmarket": ["주문번호", "장바구니번호(결제번호)", "결제번호", "주문관리번호"],
        "thirtymall": ["주문번호", "고객주문번호", "order no", "orderno"],
    },
    "품목명": {
        "coupang": ["노출상품명(옵션명)", "노출상품명", "등록상품명", "상품명"],
        "smartstore": ["상품명", "주문상품명", "옵션정보", "상품명(옵션포함)", "상품명/옵션"],
        "gmarket": ["상품명", "옵션", "추가구성", "사은품", "덤"],
        "thirtymall": ["상품명", "옵션명:옵션값", "옵션", "옵션정보"],
    },
    "기타1": {
        "coupang": ["결제액", "결제금액", "상품결제금액", "payment", "결제금"],
        "smartstore": ["결제금액", "총결제금액", "상품주문금액", "판매금액", "결제 금액", "주문금액"],
        "gmarket": ["판매금액", "판매단가", "정산예정금액"],
        "thirtymall": ["판매가(할인적용가)", "결제금액", "주문금액", "총결제금액"],
    },
    "내품수량": {
        "coupang": ["구매수", "수량", "구매수량", "qty", "수량(개)"],
        "smartstore": ["수량", "주문수량", "구매수량", "상품수량", "qty"],
        "gmarket": ["수량", "주문수량", "구매수량", "qty"],
        "thirtymall": ["수량", "주문수량", "qty"],
    },
    "받는분성명": {
        "coupang": ["수취인이름", "수취인", "받는분", "수령인", "recipient"],
        "smartstore": ["수취인명", "수취인 이름", "수취인", "수령인", "받는사람", "받는분", "수하인명"],
        "gmarket": ["수령인명", "수취인명", "수취인", "수령인", "받는분"],
        "thirtymall": ["수령자명", "수취인명", "수취인", "수령인", "받는분"],
    },
    "받는분전화번호": {
        "coupang": ["수취인연락처", "전화번호", "수취인전화번호", "휴대폰", "연락처"],
        "smartstore": [
            "수취인연락처1", "수취인연락처2", "수취인연락처", "수취인 휴대전화", "수취인휴대전화",
            "수취인전화번호", "연락처", "휴대폰번호", "휴대전화"
        ],
        "gmarket": ["수령인 휴대폰", "수령인 전화번호", "수령인연락처", "연락처", "휴대폰"],
        "thirtymall": ["수령자연락처", "연락처", "휴대폰", "전화번호"],
    },
    "받는분우편번호": {
        "coupang": ["우편번호", "수취인우편번호", "배송지우편번호", "zip", "postcode"],
        "smartstore": ["수취인우편번호", "우편번호", "배송지우편번호", "수취인 우편번호", "우편 번호"],
        "gmarket": ["우편번호", "배송지우편번호", "수취인우편번호"],
        "thirtymall": ["우편번호", "배송지우편번호", "수취인우편번호"],
    },
    "받는분주소(전체,분할)": {
        "coupang": ["주소", "수취인주소", "배송지주소", "도로명주소", "받는분주소", "주소(전체,분할)"],
        "smartstore": [
            "수취인주소", "배송지주소", "배송지", "주소",
            "수취인기본주소", "수취인상세주소", "기본주소", "상세주소",
            "도로명주소", "지번주소"
        ],
        "gmarket": ["주소", "배송지주소", "수취인주소", "도로명주소", "지번주소"],
        "thirtymall": ["주소", "배송지주소", "수취인주소", "도로명주소", "지번주소", "상세주소"],
    },
    "배송메세지1": {
        "coupang": ["배송메시지", "배송메세지", "요청사항", "배송요청사항", "message"],
        "smartstore": ["배송메시지", "배송메세지", "배송 요청사항", "배송요청사항", "배송메모", "요청사항"],
        "gmarket": ["배송시 요구사항", "배송요청사항", "배송메시지", "배송메세지", "요청사항"],
        "thirtymall": ["배송메모"],
    },
}

def build_mapping(df: pd.DataFrame, platform: str):
    mapping = {}
    for invoice_col, p_dict in CANDIDATES.items():
        if platform == "unknown":
            col = (
                find_col(df, p_dict.get("smartstore", []))
                or find_col(df, p_dict.get("coupang", []))
                or find_col(df, p_dict.get("gmarket", []))
                or find_col(df, p_dict.get("thirtymall", []))
            )
        else:
            col = find_col(df, p_dict.get(platform, []))
        mapping[invoice_col] = col
    return mapping

# -------------------------
# ✅ 스마트스토어 품목명 결합 (Q열 + S열 옵션정보)
# -------------------------
def build_smartstore_item_name(order_df: pd.DataFrame) -> pd.Series:
    product_col = find_col(order_df, ["상품명", "주문상품명", "상품명(옵션포함)", "상품명/옵션"])
    option_col = find_col(order_df, ["옵션정보", "옵션", "옵션명", "옵션내용"])

    if product_col is not None:
        product = clean_series(order_df[product_col])
    else:
        product = clean_series(order_df.iloc[:, 16]) if order_df.shape[1] > 16 else pd.Series([""] * len(order_df))

    if option_col is not None:
        option = clean_series(order_df[option_col])
    else:
        option = clean_series(order_df.iloc[:, 18]) if order_df.shape[1] > 18 else pd.Series([""] * len(order_df))

    combined = (product + " " + option).str.replace(r"\s+", " ", regex=True).str.strip()
    return combined

# -------------------------
# ✅ 쿠팡 품목명: M열 노출상품명(옵션명)
# -------------------------
def build_coupang_item_name(order_df: pd.DataFrame) -> pd.Series:
    col = find_col(order_df, ["노출상품명(옵션명)", "노출상품명", "노출 상품명(옵션명)", "노출 상품명"])
    if col is not None:
        return clean_series(order_df[col])
    if order_df.shape[1] > 12:
        return clean_series(order_df.iloc[:, 12])  # M열 fallback
    return pd.Series([""] * len(order_df))

# -------------------------
# ✅ thirtymall(떠리몰) 품목명: S열 + V열 (중복 글 1회 표기)
# -------------------------
def dedupe_merge_text(a: pd.Series, b: pd.Series) -> pd.Series:
    a = clean_series(a)
    b = clean_series(b)

    def merge_one(x, y):
        x = (x or "").strip()
        y = (y or "").strip()
        if not x and not y:
            return ""
        if not x:
            return y
        if not y:
            return x

        # 완전 동일/포함 관계면 하나만
        if x == y:
            return x
        if x in y:
            return y
        if y in x:
            return x

        # 단어(공백 기준) 중복 제거 결합
        tokens = []
        seen = set()
        for t in (x + " " + y).split():
            if t not in seen:
                seen.add(t)
                tokens.append(t)
        return " ".join(tokens)

    return pd.Series([merge_one(x, y) for x, y in zip(a.tolist(), b.tolist())])

def build_thirtymall_item_name(order_df: pd.DataFrame) -> pd.Series:
    # 헤더 기반(우선)
    s_col = find_col(order_df, ["상품명"])
    v_col = find_col(order_df, ["옵션명:옵션값", "옵션정보", "옵션", "옵션명"])

    if s_col is not None:
        s = order_df[s_col]
    else:
        s = order_df.iloc[:, 18] if order_df.shape[1] > 18 else pd.Series([""] * len(order_df))  # S열 fallback

    if v_col is not None:
        v = order_df[v_col]
    else:
        v = order_df.iloc[:, 21] if order_df.shape[1] > 21 else pd.Series([""] * len(order_df))  # V열 fallback

    return dedupe_merge_text(s, v)

# -------------------------
# ✅ gmarket(지마켓) 품목명: 상품명 + 옵션/추가구성/사은품/덤
# -------------------------
def join_nonempty_unique(values: list[str], sep: str = " / ") -> str:
    out = []
    seen = set()
    for v in values:
        v = (v or "").strip()
        if not v:
            continue
        if v in seen:
            continue
        seen.add(v)
        out.append(v)
    return sep.join(out)

def build_gmarket_item_name(order_df: pd.DataFrame) -> pd.Series:
    product_col = find_col(order_df, ["상품명", "주문상품명", "상품명(옵션포함)"])
    product = clean_series(order_df[product_col]) if product_col is not None else pd.Series([""] * len(order_df))

    option_cols = [
        find_col(order_df, ["옵션"]),
        find_col(order_df, ["추가구성"]),
        find_col(order_df, ["사은품"]),
        find_col(order_df, ["덤"]),
    ]

    parts = [product]
    for c in option_cols:
        if c is not None:
            parts.append(clean_series(order_df[c]))

    if not parts:
        return pd.Series([""] * len(order_df))

    parts_df = pd.concat(parts, axis=1)
    return parts_df.apply(lambda row: join_nonempty_unique(row.tolist()), axis=1)

def build_gmarket_phone(order_df: pd.DataFrame) -> pd.Series:
    mobile_col = find_col(order_df, ["수령인 휴대폰", "수령인휴대폰", "수령인 연락처", "휴대폰"])
    land_col = find_col(order_df, ["수령인 전화번호", "수령인전화번호", "전화번호"])

    mobile = clean_series(order_df[mobile_col]) if mobile_col is not None else pd.Series([""] * len(order_df))
    land = clean_series(order_df[land_col]) if land_col is not None else pd.Series([""] * len(order_df))

    if mobile_col is None:
        return land
    if land_col is None:
        return mobile
    return mobile.where(mobile != "", land)

def build_gmarket_pickup_date(order_df: pd.DataFrame) -> pd.Series:
    deadline_col = find_col(order_df, ["발송마감일", "발송 마감일"])
    planned_col = find_col(order_df, ["발송예정일", "발송 예정일", "집하예정일", "출고예정일"])

    deadline = clean_series(order_df[deadline_col]) if deadline_col is not None else pd.Series([""] * len(order_df))
    planned = clean_series(order_df[planned_col]) if planned_col is not None else pd.Series([""] * len(order_df))

    if deadline_col is None:
        return planned
    if planned_col is None:
        return deadline
    return deadline.where(deadline != "", planned)

def build_gmarket_item_code(order_df: pd.DataFrame) -> pd.Series:
    col = find_col(order_df, ["상품번호", "판매자 관리코드", "판매자 상세관리코드", "SKU번호", "SKU번호 및 수량"])
    if col is None:
        return pd.Series([""] * len(order_df))
    return clean_series(order_df[col])

def build_gmarket_tracking_no(order_df: pd.DataFrame) -> pd.Series:
    col = find_col(order_df, ["송장번호", "운송장번호"])
    if col is None:
        return pd.Series([""] * len(order_df))
    return clean_series(order_df[col])

# -------------------------
# ✅ thirtymall(떠리몰) 주문번호: H열(8번째 컬럼) 강제
# -------------------------
def build_thirtymall_order_no(order_df: pd.DataFrame) -> pd.Series:
    if order_df.shape[1] > 7:
        return clean_series(order_df.iloc[:, 7])
    return pd.Series([""] * len(order_df))

# -------------------------
# 스마트스토어 받는사람 보강(전화/우편/주소)
# -------------------------
def build_smartstore_phone(order_df: pd.DataFrame) -> pd.Series:
    c1 = find_col(order_df, ["수취인연락처1", "수취인연락처(1)", "수취인 휴대전화", "수취인휴대전화"])
    c2 = find_col(order_df, ["수취인연락처2", "수취인연락처(2)"])
    c  = find_col(order_df, ["수취인연락처", "수취인전화번호", "연락처", "휴대폰번호", "휴대전화"])

    if c1 is not None:
        return clean_series(order_df[c1])
    if c2 is not None:
        return clean_series(order_df[c2])
    if c is not None:
        return clean_series(order_df[c])
    return pd.Series([""] * len(order_df))

def build_smartstore_zip(order_df: pd.DataFrame) -> pd.Series:
    z = find_col(order_df, ["수취인우편번호", "우편번호", "배송지우편번호", "우편 번호"])
    if z is None:
        return pd.Series([""] * len(order_df))
    return clean_series(order_df[z])

def build_smartstore_address(order_df: pd.DataFrame) -> pd.Series:
    base = find_col(order_df, ["수취인기본주소", "기본주소", "도로명주소", "지번주소"])
    detail = find_col(order_df, ["수취인상세주소", "상세주소", "상세 주소"])

    if base is not None:
        base_s = clean_series(order_df[base])
        if detail is not None:
            detail_s = clean_series(order_df[detail])
            return (base_s + " " + detail_s).str.replace(r"\s+", " ", regex=True).str.strip()
        return base_s

    addr = find_col(order_df, ["수취인주소", "배송지주소", "배송지", "주소"])
    if addr is None:
        return pd.Series([""] * len(order_df))
    return clean_series(order_df[addr])

# -------------------------
# 송장 행 생성
# -------------------------
def make_invoice_rows(template_columns: list[str], order_df: pd.DataFrame, mapping: dict, platform: str) -> pd.DataFrame:
    out = pd.DataFrame({c: [""] * len(order_df) for c in template_columns})

    # 기본 매핑
    for inv_col, ord_col in mapping.items():
        if inv_col in out.columns and ord_col is not None and ord_col in order_df.columns:
            out[inv_col] = order_df[ord_col]

    # 플랫폼별 품목명 강제 규칙 적용
    if "품목명" in out.columns:
        if platform == "smartstore":
            out["품목명"] = build_smartstore_item_name(order_df)
        elif platform == "coupang":
            out["품목명"] = build_coupang_item_name(order_df)
        elif platform == "gmarket":
            out["품목명"] = build_gmarket_item_name(order_df)
        elif platform == "thirtymall":
            out["품목명"] = build_thirtymall_item_name(order_df)

    # 떠리몰 주문번호는 H열 강제 사용
    if platform == "thirtymall" and "고객주문번호" in out.columns:
        out["고객주문번호"] = build_thirtymall_order_no(order_df)

    # 스마트스토어 받는사람 정보 강제 세팅(분리 컬럼 조합 포함)
    if platform == "smartstore":
        if "받는분전화번호" in out.columns:
            out["받는분전화번호"] = build_smartstore_phone(order_df)
        if "받는분우편번호" in out.columns:
            out["받는분우편번호"] = build_smartstore_zip(order_df)
        if "받는분주소(전체,분할)" in out.columns:
            out["받는분주소(전체,분할)"] = build_smartstore_address(order_df)

    # 지마켓 보강(전화/집하예정일/품목코드/운송장번호)
    if platform == "gmarket":
        if "받는분전화번호" in out.columns:
            out["받는분전화번호"] = build_gmarket_phone(order_df)
        if "집하예정일" in out.columns:
            out["집하예정일"] = build_gmarket_pickup_date(order_df)
        if "품목코드" in out.columns:
            out["품목코드"] = build_gmarket_item_code(order_df)
        if "운송장번호" in out.columns:
            out["운송장번호"] = build_gmarket_tracking_no(order_df)

    return out

# =========================
# UI: 템플릿 선택
# =========================
template_mode = st.radio(
    "송장 템플릿 불러오기 방식",
    ["기본 템플릿 사용(추천)", "템플릿 파일 직접 업로드"],
    horizontal=True
)

template_upload = None
if template_mode == "템플릿 파일 직접 업로드":
    template_upload = st.file_uploader("송장 템플릿 파일 업로드 (xlsx)", type=["xlsx"], key="template")

uploaded_files = st.file_uploader(
    "주문 파일들을 업로드하세요 (xlsx) - 여러 개 선택 가능",
    type=["xlsx"],
    accept_multiple_files=True
)

platform_label = {
    "coupang": "쿠팡",
    "smartstore": "스마트스토어",
    "gmarket": "지마켓",
    "thirtymall": "thirtymall(떠리몰)",
    "unknown": "알수없음",
}

if uploaded_files:
    try:
        # 템플릿 로드
        if template_upload is not None:
            template_df = pd.read_excel(template_upload)
            template_columns = list(template_df.columns)
        else:
            template_df = build_default_template_df()
            template_columns = DEFAULT_TEMPLATE_COLUMNS

        all_out_rows = []
        report_rows = []

        for uf in uploaded_files:
            # 1) 플랫폼 판별 (스마트스토어는 암호 때문에 일반 로드 실패할 수 있음)
            try:
                tmp_df = read_excel_safely(uf, platform_hint=None)
                platform = detect_platform(tmp_df)
            except Exception:
                # 일반 로드 실패 => 스마트스토어 가능성이 높으므로 복호화+header=1로 로드 후 판별
                tmp_df2 = read_excel_safely(uf, platform_hint="smartstore")
                platform = detect_platform(tmp_df2)

            # 2) 정식 로드
            if platform == "smartstore":
                order_df = read_excel_safely(uf, platform_hint="smartstore")
            else:
                order_df = read_excel_safely(uf, platform_hint=None)

            mapping = build_mapping(order_df, platform)
            out_rows = make_invoice_rows(template_columns, order_df, mapping, platform)
            all_out_rows.append(out_rows)

            ok_cnt = sum(1 for v in mapping.values() if v is not None)
            report_rows.append({
                "파일명": uf.name,
                "자동판별 플랫폼": platform_label.get(platform, "알수없음"),
                "매핑 성공(참고)": f"{ok_cnt}/{len(mapping)}",
                "행(주문) 수": len(order_df),
            })

        merged_out = pd.concat(all_out_rows, ignore_index=True)

        st.subheader("📌 파일별 자동 판별/변환 요약")
        st.dataframe(pd.DataFrame(report_rows), use_container_width=True)

        # 상품관리코드(품목코드) 유무로 합배 / 단품 분리
        품목코드_col = "품목코드"
        if 품목코드_col in merged_out.columns:
            has_code_mask = (
                merged_out[품목코드_col]
                .astype(str)
                .str.strip()
                .replace(["nan", "None", ""], None)
                .notna()
            )
            df_합배 = merged_out[has_code_mask].reset_index(drop=True)
            df_단품 = merged_out[~has_code_mask].reset_index(drop=True)
        else:
            df_합배 = pd.DataFrame(columns=merged_out.columns)
            df_단품 = merged_out.copy()

        # 주소 오름차순 정렬
        addr_col = "받는분주소(전체,분할)"
        if addr_col in df_합배.columns:
            df_합배 = df_합배.sort_values(addr_col, na_position="last").reset_index(drop=True)
        if addr_col in df_단품.columns:
            df_단품 = df_단품.sort_values(addr_col, na_position="last").reset_index(drop=True)

        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")

        st.success(
            f"✅ 변환 완료! 총 {len(merged_out)}행  |  합배: {len(df_합배)}행  |  단품: {len(df_단품)}행"
        )

        col1, col2 = st.columns(2)

        with col1:
            buf_합배 = BytesIO()
            with pd.ExcelWriter(buf_합배, engine="openpyxl") as writer:
                df_합배.to_excel(writer, index=False)
            buf_합배.seek(0)
            st.download_button(
                f"📥 합배 송장파일 다운로드 ({len(df_합배)}행)",
                data=buf_합배.getvalue(),
                file_name=f"합배_송장파일_{now_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col2:
            buf_단품 = BytesIO()
            with pd.ExcelWriter(buf_단품, engine="openpyxl") as writer:
                df_단품.to_excel(writer, index=False)
            buf_단품.seek(0)
            st.download_button(
                f"📥 단품 송장파일 다운로드 ({len(df_단품)}행)",
                data=buf_단품.getvalue(),
                file_name=f"단품_송장파일_{now_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")
