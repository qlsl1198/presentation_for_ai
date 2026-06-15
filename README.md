# AI PPT Generator

주제를 입력하면 발표 목차와 슬라이드 초안을 만들고, `.pptx` 파일로 다운로드할 수 있는 Streamlit 앱입니다.

## 기능

- 발표 주제, 대상, 슬라이드 수, 발표 톤 입력
- AI API 키가 있으면 OpenAI로 슬라이드 초안 생성
- API 키가 없어도 기본 규칙 기반 초안 생성
- 생성된 내용을 화면에서 확인
- PowerPoint 파일 다운로드

## 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

OpenAI API를 사용하려면 실행 전에 환경 변수를 설정합니다.

```bash
export OPENAI_API_KEY="your_api_key"
streamlit run app.py
```

## 과제 주제

AI를 활용한 스타트업 서비스 기획 및 구현

- 주제: AI PPT 생성기
- 구현: 사용자가 발표 주제를 입력하면 AI가 발표 흐름과 슬라이드 내용을 생성하고 PPTX 파일로 저장
- 구현 플랫폼: Streamlit
- AI 적용: 발표 구조 생성, 슬라이드 제목 및 핵심 문장 생성
