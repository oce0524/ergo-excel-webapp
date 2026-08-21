# 엑셀 누적 데이터베이스 웹앱

판매(이카운트) / 재고(이지어드민) 엑셀을 각각 업로드해 **Supabase(PostgreSQL)** 에 누적 저장하는 Streamlit 앱입니다.

## 기능
- 판매 / 재고 엑셀을 **각각 따로** 업로드
- **캘린더로 기준일자 선택** 후 저장
- DB에 계속 **누적** (같은 날짜는 '교체' 옵션으로 중복 방지)
- **비밀번호 로그인**으로 화면 보호
- 누적 데이터 조회 · 엑셀 다운로드 · 특정 날짜 삭제

## 로컬 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```
접속 정보는 `.streamlit/secrets.toml` 에 넣습니다. (예시: `.streamlit/secrets.toml.example`)

## Streamlit Community Cloud 배포
1. 이 저장소(private)를 Streamlit Cloud 에 연결
2. **App settings > Secrets** 에 `secrets.toml` 내용을 붙여넣기
3. 메인 파일: `app.py`

## 보안
- `.streamlit/secrets.toml` 은 `.gitignore` 로 **깃허브에 올라가지 않습니다.**
- DB 접속정보/비밀번호는 코드에 없고 **secrets 로만** 읽습니다.

## 데이터베이스
- Supabase PostgreSQL, 테이블 `sales`(판매), `stock`(재고)
- 각 행에 `기준일자`, `업로드시각`, `원본파일` 메타컬럼이 추가됩니다.
- IPv4 환경/클라우드 호환을 위해 **Connection Pooler** 주소를 사용합니다.
