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
import json
from datetime import date, datetime

import requests
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


# ------------------------------------------------------------------ 설정(config) 저장
def ensure_config_table():
    with get_engine().begin() as c:
        c.execute(text('CREATE TABLE IF NOT EXISTS app_config '
                       '(config_key text PRIMARY KEY, config_value text)'))


def get_config(key, default=""):
    try:
        ensure_config_table()
        with get_engine().connect() as c:
            row = c.execute(text('SELECT config_value FROM app_config WHERE config_key=:k'),
                            {"k": key}).fetchone()
        return row[0] if row and row[0] is not None else default
    except Exception:
        return default


def set_config(key, value):
    ensure_config_table()
    with get_engine().begin() as c:
        c.execute(text('INSERT INTO app_config(config_key, config_value) VALUES(:k, :v) '
                       'ON CONFLICT (config_key) DO UPDATE SET config_value=:v'),
                  {"k": key, "v": value})


# ------------------------------------------------------------------ 이지어드민 수집
import re as _re
import html as _html
import time as _time
from urllib.parse import urlencode as _urlencode

EZ_URL = "https://ga25.ezadmin.co.kr/function.htm"
EZ_TABLE = "ezorders"          # 이지어드민 주문 누적 테이블
EZ_MAX_PAGES = 500             # 안전 상한


def _ez_headers(cookie):
    return {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151 Safari/537.36",
        "Cookie": cookie,
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Origin": "https://ga25.ezadmin.co.kr",
        "Referer": "https://ga25.ezadmin.co.kr/template35.htm?template=DS00",
        "Accept": "application/json, text/javascript, */*; q=0.01",
    }


def _ez_inner_par(d, page):
    """이지어드민 DS00 주문 그리드의 검색조건(par). 날짜만 바꿔 재현."""
    return [
        ("template", "DS00"), ("action", ""), ("search", "1"), ("page", str(page)),
        ("_sort", ""), ("sort_order", ""), ("panel_open", "false"), ("field_change", ""),
        ("bck_search", "0"), ("recover_delete", ""),
        ("date_type", "collect_date"),
        ("start_date", d), ("start_hour", "00:00:00"),
        ("end_date", d), ("end_hour", "23:59:59"),
        ("date_period_sel", "2"),
        ("option[]", "seq"), ("query_str[]", ""),
        ("multi_supply_group", ""), ("multi_supply", ""), ("str_supply_code", "0"),
        ("status_sel", "0"), ("pack_sel", "0"), ("check_set_match", "0"),
        ("order_cs_sel", "0"), ("work_type", "0"), ("checkbox_options_string", ""),
        ("trans_corp", "99"), ("user_area", ""),
        ("multi_shop_group", ""), ("multi_shop", ""), ("str_shop_code", "0"),
        ("tags_string", ""), ("product_tag_include_type", "1"),
        ("labels_string", ""), ("order_label_include_type", "1"),
        ("date_type2", "0"),
        ("start_date2", d), ("start_hour2", "00"),
        ("end_date2", d), ("end_hour2", "23"),
        ("date_period_sel2", "0"), ("category", "0"),
        ("option[]", ""), ("query_str[]", ""),
        ("c_cs", "blink"), ("order_copy", "0"), ("create_order", "0"),
        ("print_enable", "0"), ("product_expect", "0"),
        ("return_money_expect_price", ""), ("return_money_return_price", ""),
        ("trans_who", "0"), ("cs_reason", ""), ("multi_user_cs_type", ""),
        ("user_cs_type", "0"), ("trans_type", ""),
        ("special_option[]", "사은품선택"),
        ("select_field", "DS00"), ("download_field", "DS00_file"), ("download_type", "0"),
    ]


class EzCookieExpired(Exception):
    pass


def _ez_fetch(date_str, page, cookie, sess, rows=100):
    par = _urlencode(_ez_inner_par(date_str, page))
    body = [
        ("_search", "false"), ("nd", str(int(_time.time() * 1000))),
        ("rows", str(rows)), ("page", str(page)), ("sidx", ""), ("sord", "asc"),
        ("template", "DS00"), ("action", "grid_DS00"), ("bck_search", "0"),
        ("par", par),
    ]
    r = sess.post(EZ_URL, headers=_ez_headers(cookie), data=_urlencode(body), timeout=30)
    txt = r.text.lstrip()
    if txt.startswith("<script") or "location.href" in txt[:60] or "'/index.html'" in txt[:80]:
        raise EzCookieExpired("쿠키가 만료되었거나 유효하지 않습니다. 다시 입력해 주세요.")
    return json.loads(r.text, strict=False)


def _ez_clean(v):
    if not isinstance(v, str):
        return v
    t = _re.sub(r"<[^>]+>", "", v)
    t = _html.unescape(t)
    return _re.sub(r"\s+", " ", t).strip()


def ez_collect(date_str, cookie, progress=None):
    """지정 날짜(수집일)의 모든 주문을 전 페이지 순회해 정리된 dict 목록으로 반환."""
    sess = requests.Session()
    first = _ez_fetch(date_str, 1, cookie, sess)
    pages = min(int(first.get("total", 1) or 1), EZ_MAX_PAGES)
    rows = list(first.get("rows", []))
    if progress:
        progress(1, pages, len(rows))
    for p in range(2, pages + 1):
        data = _ez_fetch(date_str, p, cookie, sess)
        rows += data.get("rows", [])
        if progress:
            progress(p, pages, len(rows))
        _time.sleep(0.15)
    # cell 추출 + HTML 정리 (상품라인 그대로 유지, 인메모리 중복제거 안 함)
    recs = []
    for r in rows:
        cell = {k: _ez_clean(val) for k, val in r.get("cell", {}).items()}
        cell.pop("chk", None)   # 체크박스 위젯 컬럼 제거
        recs.append(cell)
    return recs


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


# ------------------------------------------------------------------ 이지어드민 수집 탭
def ezadmin_section():
    st.subheader("🛒 이지어드민 주문 수집")
    st.caption("이지어드민(DS00 주문) 데이터를 날짜별로 가져와 DB에 누적합니다. (엑셀 업로드 없이 바로 저장)")

    saved_cookie = get_config("ezadmin_cookie", "")
    with st.expander("🔐 쿠키 설정", expanded=not bool(saved_cookie)):
        st.write("Chrome 개발자도구(F12) → Network 탭 → `function.htm` 요청의 **Cookie 값**을 붙여넣으세요. "
                 "세션이 만료되면 다시 넣어야 합니다.")
        cookie_in = st.text_area("이지어드민 쿠키", value=saved_cookie, height=110, key="ez_cookie")
        if st.button("쿠키 저장", key="ez_save_cookie"):
            set_config("ezadmin_cookie", cookie_in.strip())
            st.success("쿠키를 저장했습니다. (DB에 보관 · 다음부터 자동 사용)")
        if saved_cookie:
            st.caption("✔ 저장된 쿠키가 있습니다.")

    col1, col2 = st.columns([1, 2])
    with col1:
        d = st.date_input("📅 수집할 날짜 (수집일 기준)", value=date.today(),
                          format="YYYY-MM-DD", key="ez_date")
    with col2:
        st.write(" ")
        run = st.button("⬇ 이 날짜 주문 수집", type="primary", key="ez_run")

    if run:
        cookie = (st.session_state.get("ez_cookie") or saved_cookie or "").strip()
        if not cookie:
            st.error("먼저 쿠키를 입력/저장하세요.")
        else:
            prog = st.progress(0.0)
            status = st.empty()

            def cb(p, total, n):
                prog.progress(min(1.0, p / max(total, 1)))
                status.write(f"수집 중… {p}/{total} 페이지 · 누적 {n:,}행")

            recs = None
            try:
                recs = ez_collect(str(d), cookie, cb)
            except EzCookieExpired as e:
                st.error(f"❌ {e}")
            except Exception as e:
                st.error(f"수집 실패: {e}")

            if recs is not None:
                if not recs:
                    st.warning(f"{d} 에 해당하는 주문이 없습니다.")
                else:
                    df = pd.DataFrame(recs)
                    df.insert(0, "기준일자", str(d))
                    df["수집시각"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    try:
                        delete_by_date(EZ_TABLE, str(d))   # 같은 날짜 재수집 시 중복 방지
                        append_df(EZ_TABLE, df)
                        prog.progress(1.0)
                        st.success(f"✅ 수집 완료! {len(df):,}건 저장 (기준일자 {d})")
                    except Exception as e:
                        st.error(f"저장 실패: {e}")

    st.markdown("**📊 수집 누적 현황**")
    try:
        data = read_table(EZ_TABLE)
    except Exception as e:
        st.error(f"DB 조회 실패: {e}")
        data = pd.DataFrame()
    if data.empty:
        st.info("아직 수집된 주문이 없습니다.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("누적 주문(행) 수", f"{len(data):,}")
        if "기준일자" in data.columns:
            c2.metric("수집한 날짜 수", f"{data['기준일자'].nunique()}")
            by_date = (data.groupby("기준일자").size().reset_index(name="행수").sort_values("기준일자"))
            st.dataframe(by_date, use_container_width=True, hide_index=True)
            with st.expander("🗑 특정 날짜 삭제"):
                tgt = st.selectbox("삭제할 기준일자", by_date["기준일자"].tolist(), key="ez_del")
                if st.button("이 날짜 데이터 삭제", key="ez_delbtn"):
                    delete_by_date(EZ_TABLE, tgt)
                    st.warning(f"'{tgt}' 삭제 완료. 새로고침(R)하세요.")
        with st.expander("🔎 수집 데이터 보기 / 다운로드"):
            st.dataframe(data.head(500), use_container_width=True, height=360)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                data.to_excel(w, index=False, sheet_name="ezorders")
            st.download_button("⬇ 수집 데이터 엑셀로 내려받기", data=buf.getvalue(),
                               file_name="이지어드민_주문_누적.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="ez_dl")


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

tab1, tab2, tab3, tab4 = st.tabs(
    ["📈 판매 업로드", "📦 재고 업로드", "🛒 이지어드민 수집", "🗄 전체 요약"])
with tab1:
    upload_section("sales")
with tab2:
    upload_section("stock")
with tab3:
    ezadmin_section()
with tab4:
    st.subheader("전체 누적 요약")
    summary = list(KINDS.items()) + [("ez", {"label": "🛒 이지어드민 주문", "table": EZ_TABLE})]
    for kind_key, cfg in summary:
        try:
            data = read_table(cfg["table"])
            rows = len(data)
            dates = data["기준일자"].nunique() if ("기준일자" in data.columns and not data.empty) else 0
        except Exception:
            rows, dates = 0, 0
        st.write(f"- **{cfg['label']}** (`{cfg['table']}` 테이블): 누적 {rows:,}행 / 기준일자 {dates}개")
