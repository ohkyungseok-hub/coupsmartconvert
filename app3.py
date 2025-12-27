import re
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="주문파일 → 송장파일 변환", layout="centered")
st.title("📦 주문파일 → 송장 출력용 파일 변환기 (자동 플랫폼 판별 + 다중 업로드 통합)")

st.markdown("""
- 쿠팡/스마트스토어 주문 엑셀(xlsx) **여러 개를 한번에 업로드**
- 파일별로 **플랫폼 자동 판별**
- **헤더(컬럼명) 기반 자동 매핑**
- 결과는 **한 개의 송장파일로 통합 변환**
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
    # 빈 템플릿(컬럼만 존재). 필요하면 기본값도 여기서 세팅 가능.
    return pd.DataFrame(columns=DEFAULT_TEMPLATE_COLUMNS)

# -------------------------
# 유틸: 컬럼명 정규화
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

# -------------------------
# 플랫폼 판별용 시그니처(헤더 키워드)
# -------------------------
PLATFORM_SIGNATURES = {
    "coupang": ["등록상품명", "수취인이름", "주문번호", "결제액", "구매수", "배송메시지", "배송메세지"],
    "smartstore": ["상품주문번호", "수취인명", "배송메시지", "배송메세지", "옵션정보", "주문번호", "우편번호"],
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

    if coupang_score == 0 and smart_score == 0:
        return "unknown"
    return "coupang" if coupang_score >= smart_score else "smartstore"

# -------------------------
# 송장필드(템플릿 컬럼명)별 후보 헤더명(자동 매핑)
# ※ 템플릿 컬럼명과 동일하게 맞춤!
# -------------------------
CANDIDATES = {
    "고객주문번호": {
        "coupang": ["주문번호", "고객주문번호", "order number", "orderno"],
        "smartstore": ["주문번호", "상품주문번호", "상품 주문번호", "주문관리번호", "order no"],
    },
    "품목명": {
        "coupang": ["등록상품명", "상품명", "옵션정보", "product name"],
        "smartstore": ["상품명", "옵션정보", "상품명(옵션포함)", "상품명/옵션", "주문상품명"],
    },
    "기타1": {
        "coupang": ["결제액", "결제금액", "상품결제금액", "payment", "결제금"],
        "smartstore": ["결제금액", "상품주문금액", "총결제금액", "판매금액", "결제 금액"],
    },
    "내품수량": {
        "coupang": ["구매수", "수량", "구매수량", "qty", "수량(개)"],
        "smartstore": ["수량", "구매수량", "주문수량", "상품수량", "qty"],
    },
    "받는분성명": {
        "coupang": ["수취인이름", "수취인", "받는분", "수령인", "recipient"],
        "smartstore": ["수취인명", "수취인", "수령인", "받는사람", "받는분", "수취인 이름"],
    },
    "받는분전화번호": {
        "coupang": ["수취인연락처", "전화번호", "수취인전화번호", "휴대폰", "연락처"],
        "smartstore": ["수취인연락처1", "수취인연락처", "수취인연락처(1)", "수취인 휴대전화", "수취인전화번호", "연락처"],
    },
    "받는분우편번호": {
        "coupang": ["우편번호", "수취인우편번호", "배송지우편번호", "zip", "postcode"],
        "smartstore": ["우편번호", "수취인우편번호", "배송지우편번호", "수취인 우편번호"],
    },
    "받는분주소(전체,분할)": {
        "coupang": ["주소", "수취인주소", "배송지주소", "도로명주소", "받는분주소", "주소(전체,분할)"],
        "smartstore": ["배송지", "배송지주소", "수취인주소", "기본주소", "도로명주소", "주소"],
    },
    "배송메세지1": {
        "coupang": ["배송메시지", "배송메세지", "요청사항", "배송요청사항", "message"],
        "smartstore": ["배송메시지", "배송메세지", "배송 요청사항", "배송요청사항", "배송메모", "요청사항"],
    },
}

def build_mapping(df: pd.DataFrame, platform: str):
    mapping = {}
    for invoice_col, p_dict in CANDIDATES.items():
        if platform == "unknown":
            col = find_col(df, p_dict["smartstore"]) or find_col(df, p_dict["coupang"])
        else:
            col = find_col(df, p_dict[platform])
        mapping[invoice_col] = col
    return mapping

def make_invoice_rows(template_columns: list[str], order_df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    out = pd.DataFrame({c: [""] * len(order_df) for c in template_columns})
    for inv_col, ord_col in mapping.items():
        if inv_col in out.columns and ord_col is not None and ord_col in order_df.columns:
            out[inv_col] = order_df[ord_col]
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

            out_rows = make_invoice_rows(template_columns, order_df, mapping)
            all_out_rows.append(out_rows)

            ok_cnt = sum(1 for v in mapping.values() if v is not None)
            report_rows.append({
                "파일명": uf.name,
                "자동판별 플랫폼": "쿠팡" if platform == "coupang" else ("스마트스토어" if platform == "smartstore" else "알수없음"),
                "매핑 성공": f"{ok_cnt}/{len(mapping)}",
                "행(주문) 수": len(order_df),
            })

        merged_out = pd.concat(all_out_rows, ignore_index=True)

        st.subheader("📌 파일별 자동 판별/변환 요약")
        st.dataframe(pd.DataFrame(report_rows), use_container_width=True)

        # 엑셀을 메모리로 저장(서버/클라우드 배포에 유리)
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
