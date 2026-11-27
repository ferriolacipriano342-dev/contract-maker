# 리눅스 서버 배포 시 한글 깨짐 방지 폰트 안내

Streamlit Cloud, Ubuntu/Debian Docker 등 리눅스 서버에서 WeasyPrint 또는 헤드리스 PDF 변환을 실행할 때 한글 폰트가 설치되어 있지 않으면 글자가 네모(□)로 깨질 수 있습니다.

## 1. Streamlit Cloud 배포 시
프로젝트 루트의 `packages.txt` 파일에 이미 아래 패키지가 등록되어 있어 배포 시 자동 설치됩니다.

```text
fonts-nanum
fonts-noto-cjk
```

## 2. Ubuntu / Debian Linux 서버 직접 설치 시
```bash
sudo apt-get update
sudo apt-get install -y fonts-nanum fonts-noto-cjk fontconfig
fc-cache -fv
```

## 3. 커스텀 폰트(.ttf / .otf) 직접 포함 시
1. `NanumGothic.ttf` 또는 `NotoSansKR-Regular.otf` 파일을 이 `fonts/` 폴더에 복사합니다.
2. `templates/html/housing_jeonse.html`의 `<style>` 상단에 `@font-face`를 추가하면 서버 시스템 폰트 없이도 일관된 한글 PDF가 생성됩니다.

```css
@font-face {
  font-family: 'ContractKorean';
  src: url('fonts/NanumGothic.ttf') format('truetype');
  font-weight: normal;
  font-style: normal;
}
```
