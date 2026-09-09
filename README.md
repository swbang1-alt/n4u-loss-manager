# 엔포유 소분 작업 및 감모 관리

`엔포유소분리스트.xlsx`와 `제품조회.xlsx`를 업로드해 전표별 투입·산출·감모 현황을 계산하는 Streamlit 앱입니다.

## 무료 웹 배포 (Streamlit Community Cloud)

1. GitHub에서 빈 **Public** 저장소를 만듭니다. 예: `n4u-loss-manager`
2. 이 폴더에서 아래 네 파일만 새 저장소에 올립니다.
   - `n4u_upload_app.py`
   - `requirements.txt`
   - `README.md`
   - `.gitignore`
3. [Streamlit Community Cloud](https://share.streamlit.io/)에 GitHub 계정으로 로그인합니다.
4. **Create app**을 누르고 방금 만든 저장소, 브랜치, `n4u_upload_app.py`를 선택합니다.
5. 원하는 앱 주소를 정한 뒤 **Deploy**를 누릅니다.

배포 후에는 `https://원하는이름.streamlit.app`과 같은 주소로 접속합니다.

## 주의

이 앱은 사용자가 업로드하는 엑셀 파일을 분석하기 위해 서버로 전송합니다. 공개 앱 주소를 아는 누구나 파일을 업로드할 수 있으므로, 민감한 거래 데이터라면 공개 배포 전에 사내 보안 정책을 확인하세요.
