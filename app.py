# -*- coding: utf-8 -*-
"""
엑셀 업로드 → Supabase(PostgreSQL) 누적 웹앱 (Streamlit)
- 판매(이카운트) / 재고(이지어드민) 두 종류를 각각 따로 업로드
- 캘린더로 '기준일자' 선택 후 저장
- Supabase(PostgreSQL) 데이터베이스에 계속 누적
- 접속 정보는 코드가 아닌 st.secrets 에서만 읽음
- 비밀번호를 입력해야만 화면이 보임
"""

import io
from datetime import date, datetime

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text, inspect
from sqlalchemy.engine import URL

# ------------------------------------------------------------------ 기본 설정
st.set_page_config(page_title="엑셀 누적 데이터베이스", page_icon="🗄", layout="wide")

KINDS = {
    "sales": {"label": "📈 판매 (이카운트)", "table": "sales",
              "must_have": ["일자", "판매금액"]},
    "stock": {"label": "📦 재고 (이지어드민)", "table": "stock",
              "must_have": ["상품명", "정상재고"]},
}


# ------------------------------------------------------------------ 비밀번호 잠금
def require_password():
    """secrets 의 app_password 와 일치해야 통과."""
    if st.session_state.get("authenticated"):
        return

    st.title("🔒 로그인")
    st.caption("이 앱은 비밀번호로 보호되어 있습니다.")

    with st.form("login_form", clear_on_submit=False):
        pw = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("로그인")

    if submitted:
        if pw == st.secrets.get("app_password"):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    st.stop()


# ------------------------------------------------------------------ DB
@st.cache_resource
def get_engine():
    if "postgres" not in st.secrets:
        st.error("secrets 에 [postgres] 접속 정보가 없습니다. .streamlit/secrets.toml 을 확인하세요.")
        st.stop()
    cfg = st.secrets["postgres"]
    url = URL.create(
        "postgresql+psycopg2",
        username=cfg["user"], password=cfg["password"],
        host=cfg["host"], port=int(cfg["port"]), database=cfg["dbname"],
    )
    return create_engine(url, pool_pre_ping=True, connect_args={"connect_timeout": 15})


def table_exists(table):
    try:
        return inspect(get_engine()).has_table(table)
    except Exception:
        return False


def read_table(table):
    if not table_exists(table):
        return pd.DataFrame()
    return pd.read_sql(f'SELECT * FROM "{table}"', get_engine())


def delete_by_date(table, 기준일자):
    if table_exists(table):
        with get_engine().begin() as c:
            c.execute(text(f'DELETE FROM "{table}" WHERE "기준일자"=:d'), {"d": str(기준일자)})


def append_df(table, df):
    df.to_sql(table, get_engine(), if_exists="append", index=False)


# ------------------------------------------------------------------ 엑셀 파싱
def smart_read_excel(file, must_have):
    """헤더 행 위치를 자동 탐지해서 DataFrame 반환."""
    raw = pd.read_excel(file, sheet_name=0, header=None, dtype=object)
    header_row = None
    for i in range(min(8, len(raw))):
        cells = [str(c).strip() for c in raw.iloc[i].tolist() if c is not None]
        if all(any(key == c for c in cells) for key in must_have):
            header_row = i
            break
    if header_row is None:
        header_row = 0

    df = pd.read_excel(file, sheet_name=0, header=header_row).dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, [c for c in df.columns if not c.startswith("Unnamed")]]
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(lambda v: v.strip() if isinstance(v, str) else v)
    return df


# ------------------------------------------------------------------ 업로드 섹션
def upload_section(kind_key):
    cfg = KINDS[kind_key]
    table = cfg["table"]
    st.subheader(cfg["label"])

    col1, col2 = st.columns([1, 2])
    with col1:
        기준일자 = st.date_input("📅 기준일자 (캘린더에서 선택)", value=date.today(),
                              format="YYYY-MM-DD", key=f"date_{kind_key}")
        mode = st.radio("저장 방식", ["누적 추가", "같은 기준일자 데이터는 교체"],
                        key=f"mode_{kind_key}",
                        help="'교체'는 같은 날짜로 저장된 기존 데이터를 지우고 새로 넣습니다(중복 방지).")
    with col2:
        up = st.file_uploader(f"{cfg['label']} 엑셀 업로드 (.xlsx)",
                              type=["xlsx", "xls"], key=f"file_{kind_key}")

    if up is not None:
        try:
            df = smart_read_excel(up, cfg["must_have"])
        except Exception as e:
            st.error(f"엑셀을 읽지 못했습니다: {e}")
            return

        st.write(f"미리보기 — **{len(df):,}행 × {len(df.columns)}열**")
        st.dataframe(df.head(20), use_container_width=True, height=280)

        if st.button(f"💾 '{기준일자}' 로 저장(누적)", key=f"save_{kind_key}", type="primary"):
            to_save = df.copy()
            to_save.insert(0, "기준일자", str(기준일자))
            to_save["업로드시각"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            to_save["원본파일"] = up.name
            try:
                if mode.startswith("같은"):
                    delete_by_date(table, str(기준일자))
                append_df(table, to_save)
                st.success(f"저장 완료! {len(to_save):,}행이 누적되었습니다. (기준일자 {기준일자})")
            except Exception as e:
                st.error(f"저장 실패: {e}")

    # 누적 현황 ----------------------------------------------------
    st.markdown("**📊 누적 현황**")
    try:
        data = read_table(table)
    except Exception as e:
        st.error(f"DB 조회 실패: {e}")
        return

    if data.empty:
        st.info("아직 저장된 데이터가 없습니다.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("누적 총 행 수", f"{len(data):,}")
        if "기준일자" in data.columns:
            c2.metric("저장된 기준일자 수", f"{data['기준일자'].nunique()}")
            by_date = (data.groupby("기준일자").size()
                       .reset_index(name="행수").sort_values("기준일자"))
            st.dataframe(by_date, use_container_width=True, hide_index=True)

            with st.expander("🗑 특정 기준일자 삭제"):
                target = st.selectbox("삭제할 기준일자", by_date["기준일자"].tolist(),
                                      key=f"del_{kind_key}")
                if st.button("이 날짜 데이터 삭제", key=f"delbtn_{kind_key}"):
                    delete_by_date(table, target)
                    st.warning(f"'{target}' 데이터를 삭제했습니다. 새로고침(R)하세요.")

        with st.expander("🔎 전체 누적 데이터 보기 / 다운로드"):
            st.dataframe(data, use_container_width=True, height=360)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                data.to_excel(w, index=False, sheet_name=table)
            st.download_button("⬇ 누적 데이터 엑셀로 내려받기", data=buf.getvalue(),
                               file_name=f"{table}_누적.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key=f"dl_{kind_key}")


# ------------------------------------------------------------------ 메인
require_password()   # 비밀번호 통과해야 아래 화면 표시

st.title("🗄 엑셀 업로드 · 누적 데이터베이스")
st.caption("저장소: Supabase (PostgreSQL) · 접속정보는 secrets 로만 관리")

with st.sidebar:
    st.header("설정")
    if st.button("🚪 로그아웃"):
        st.session_state["authenticated"] = False
        st.rerun()
    st.caption("판매/재고 엑셀을 각 탭에서 따로 올리면 DB에 누적됩니다.")

tab1, tab2, tab3 = st.tabs(["📈 판매 업로드", "📦 재고 업로드", "🗄 전체 요약"])
with tab1:
    upload_section("sales")
with tab2:
    upload_section("stock")
with tab3:
    st.subheader("전체 누적 요약")
    for kind_key, cfg in KINDS.items():
        try:
            data = read_table(cfg["table"])
            rows = len(data)
            dates = data["기준일자"].nunique() if ("기준일자" in data.columns and not data.empty) else 0
        except Exception:
            rows, dates = 0, 0
        st.write(f"- **{cfg['label']}** (`{cfg['table']}` 테이블): 누적 {rows:,}행 / 기준일자 {dates}개")
