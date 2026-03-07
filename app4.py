import re
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="주문파일 → 송장파일 변환", layout="centered")
st.title("📦 주문파일 → 송장 출력용 파일 변환기 (자동 플랫폼 판별 + 다중 업로드 통합)")

st.markdown("""
- 쿠팡/스마트스토어/지마켓(Gmarket) 주문 엑셀(xlsx) **여러 개를 한번에 업로드**
- 파일별로 **플랫폼 자동 판별**
- **헤더(컬럼명) 기반 자동 매핑**
- 결과는 **한 개의 송장파일로 통합 변환**
- ✅ 스마트스토어 품목명: **Q열(상품명) + S열(옵션정보)**
- ✅ 쿠팡 품목명: **M열 노출상품명(옵션명)**
- ✅ 지마켓 품목명: **상품명 + 옵션/추가구성/사은품/덤**
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
    s = re.sub(r"[()\-_/.,·]", "", s)
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

# -------------------------
# 플랫폼 판별
# -------------------------
PLATFORM_SIGNATURES = {
    "coupang": ["노출상품명", "노출상품명(옵션명)", "등록상품명", "수취인이름", "주문번호", "결제액", "구매수"],
    "smartstore": ["상품주문번호", "수취인명", "배송메시지", "배송메세지", "옵션정보", "우편번호"],
    "gmarket": [
        "판매아이디", "주문번호", "상품명", "수령인명", "수령인 휴대폰", "우편번호",
        "주소", "배송시 요구사항", "발송마감일", "택배사명(발송방법)"
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

    if coupang_score == 0 and smart_score == 0 and gmarket_score == 0:
        return "unknown"
    scores = {"coupang": coupang_score, "smartstore": smart_score, "gmarket": gmarket_score}
    return max(scores, key=scores.get)

# -------------------------
# 자동 매핑 후보(템플릿 컬럼명 기준)
# -------------------------
CANDIDATES = {
    "고객주문번호": {
        "coupang": ["주문번호", "고객주문번호", "order number", "orderno"],
        "smartstore": ["상품주문번호", "상품 주문번호", "주문번호", "주문관리번호", "order no"],
        "gmarket": ["주문번호", "장바구니번호(결제번호)", "결제번호", "주문관리번호"],
    },
    "품목명": {
        # 쿠팡 품목명은 최종적으로 "M열 노출상품명(옵션명)"으로 덮어쓰기 할 것이지만,
        # 그래도 자동 매핑 후보는 넓게 둡니다.
        "coupang": ["노출상품명(옵션명)", "노출상품명", "등록상품명", "상품명"],
        "smartstore": ["상품명", "주문상품명", "옵션정보", "상품명(옵션포함)", "상품명/옵션"],
        "gmarket": ["상품명", "옵션", "추가구성", "사은품", "덤"],
    },
    "기타1": {
        "coupang": ["결제액", "결제금액", "상품결제금액", "payment", "결제금"],
        "smartstore": ["결제금액", "총결제금액", "상품주문금액", "판매금액", "결제 금액", "주문금액"],
        "gmarket": ["판매금액", "판매단가", "정산예정금액"],
    },
    "내품수량": {
        "coupang": ["구매수", "수량", "구매수량", "qty", "수량(개)"],
        "smartstore": ["수량", "주문수량", "구매수량", "상품수량", "qty"],
        "gmarket": ["수량", "주문수량", "구매수량", "qty"],
    },
    "받는분성명": {
        "coupang": ["수취인이름", "수취인", "받는분", "수령인", "recipient"],
        "smartstore": ["수취인명", "수취인 이름", "수취인", "수령인", "받는사람", "받는분", "수하인명"],
        "gmarket": ["수령인명", "수취인명", "수취인", "수령인", "받는분"],
    },
    "받는분전화번호": {
        "coupang": ["수취인연락처", "전화번호", "수취인전화번호", "휴대폰", "연락처"],
        "smartstore": [
            "수취인연락처1", "수취인연락처2", "수취인연락처", "수취인 휴대전화", "수취인휴대전화",
            "수취인전화번호", "연락처", "휴대폰번호", "휴대전화"
        ],
        "gmarket": ["수령인 휴대폰", "수령인 전화번호", "수령인연락처", "연락처", "휴대폰"],
    },
    "받는분우편번호": {
        "coupang": ["우편번호", "수취인우편번호", "배송지우편번호", "zip", "postcode"],
        "smartstore": ["수취인우편번호", "우편번호", "배송지우편번호", "수취인 우편번호", "우편 번호"],
        "gmarket": ["우편번호", "배송지우편번호", "수취인우편번호"],
    },
    "받는분주소(전체,분할)": {
        "coupang": ["주소", "수취인주소", "배송지주소", "도로명주소", "받는분주소", "주소(전체,분할)"],
        "smartstore": [
            "수취인주소", "배송지주소", "배송지", "주소",
            "수취인기본주소", "수취인상세주소", "기본주소", "상세주소",
            "도로명주소", "지번주소"
        ],
        "gmarket": ["주소", "배송지주소", "수취인주소", "도로명주소", "지번주소"],
    },
    "배송메세지1": {
        "coupang": ["배송메시지", "배송메세지", "요청사항", "배송요청사항", "message"],
        "smartstore": ["배송메시지", "배송메세지", "배송 요청사항", "배송요청사항", "배송메모", "요청사항"],
        "gmarket": ["배송시 요구사항", "배송요청사항", "배송메시지", "배송메세지", "요청사항"],
    },
}

def build_mapping(df: pd.DataFrame, platform: str):
    mapping = {}
    for invoice_col, p_dict in CANDIDATES.items():
        if platform == "unknown":
            col = (
                find_col(df, p_dict["smartstore"])
                or find_col(df, p_dict["coupang"])
                or find_col(df, p_dict["gmarket"])
            )
        else:
            col = find_col(df, p_dict[platform])
        mapping[invoice_col] = col
    return mapping

# -------------------------
# ✅ 스마트스토어 품목명 결합 (Q열 + S열 옵션정보)
# -------------------------
def build_smartstore_item_name(order_df: pd.DataFrame) -> pd.Series:
    """
    스마트스토어 품목명 = Q열(상품명) + S열(옵션정보)
    - 헤더 기반 우선 탐색
    - 실패 시 열 위치 fallback: Q=iloc[16], S=iloc[18]
    """
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
    """
    쿠팡 품목명 = M열 '노출상품명(옵션명)' 사용
    - 헤더 기반 우선 탐색
    - 실패 시 열 위치 fallback: M=iloc[12] (A=0 → M=12)
    """
    col = find_col(order_df, ["노출상품명(옵션명)", "노출상품명", "노출 상품명(옵션명)", "노출 상품명"])
    if col is not None:
        return clean_series(order_df[col])

    # fallback: M열
    if order_df.shape[1] > 12:
        return clean_series(order_df.iloc[:, 12])

    return pd.Series([""] * len(order_df))

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

    # ✅ 플랫폼별 품목명 강제 규칙 적용
    if "품목명" in out.columns:
        if platform == "smartstore":
            out["품목명"] = build_smartstore_item_name(order_df)
        elif platform == "coupang":
            out["품목명"] = build_coupang_item_name(order_df)
        elif platform == "gmarket":
            out["품목명"] = build_gmarket_item_name(order_df)

    # ✅ 스마트스토어 받는사람 정보 강제 세팅(분리 컬럼 조합 포함)
    if platform == "smartstore":
        if "받는분전화번호" in out.columns:
            out["받는분전화번호"] = build_smartstore_phone(order_df)
        if "받는분우편번호" in out.columns:
            out["받는분우편번호"] = build_smartstore_zip(order_df)
        if "받는분주소(전체,분할)" in out.columns:
            out["받는분주소(전체,분할)"] = build_smartstore_address(order_df)

    if platform == "gmarket":
        if "받는분전화번호" in out.columns:
            out["받는분전화번호"] = build_gmarket_phone(order_df)

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
            order_df = pd.read_excel(uf)

            platform = detect_platform(order_df)
            mapping = build_mapping(order_df, platform)

            out_rows = make_invoice_rows(template_columns, order_df, mapping, platform)
            all_out_rows.append(out_rows)

            ok_cnt = sum(1 for v in mapping.values() if v is not None)
            platform_label = {
                "coupang": "쿠팡",
                "smartstore": "스마트스토어",
                "gmarket": "지마켓",
                "unknown": "알수없음",
            }
            report_rows.append({
                "파일명": uf.name,
                "자동판별 플랫폼": platform_label.get(platform, "알수없음"),
                "매핑 성공(참고)": f"{ok_cnt}/{len(mapping)}",
                "행(주문) 수": len(order_df),
            })

        merged_out = pd.concat(all_out_rows, ignore_index=True)

        st.subheader("📌 파일별 자동 판별/변환 요약")
        st.dataframe(pd.DataFrame(report_rows), use_container_width=True)

        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"통합_송장파일_{now_str}.xlsx"

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            merged_out.to_excel(writer, index=False)
        buffer.seek(0)

        st.success(f"✅ 통합 송장파일 생성 완료! (총 {len(merged_out)}행)")
        st.download_button(
            "📥 통합 송장파일 다운로드",
            data=buffer.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")
