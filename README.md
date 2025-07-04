# 🚀 Excel Insight AI (엑셀 인사이트 AI)

엑셀 파일을 업로드하고 자연어로 질문하면, AI가 엑셀 데이터를 분석해 답변해주는 FastAPI 기반 서비스입니다. 

> **본 프로젝트는 직접 개발하였으며, 데이터 분석과 AI 활용에 관심 있는 분들을 위해 제작되었습니다.**

---

## 🌟 주요 기능

- **엑셀 업로드 & 자연어 Q&A**: 엑셀(.xlsx) 파일을 업로드하고, 궁금한 점을 한국어로 질문하면 AI가 답변!
- **LangChain + Ollama**: 최신 LLM(예: Llama3)과 임베딩을 활용한 강력한 문맥 이해
- **FAISS 벡터DB**: 대용량 데이터도 빠르게 검색
- **마크다운 변환**: 엑셀 데이터를 마크다운 표로 변환해 LLM이 쉽게 이해하도록 처리

---

## 🛠️ 설치 및 실행 방법

1. **필수 패키지 설치**
   ```bash
   pip install fastapi uvicorn langchain langchain_ollama langchain_community pandas faiss-cpu
   ```

2. **Ollama 설치 및 모델 준비**
   - [Ollama 공식 사이트](https://ollama.com/)에서 설치
   - 예시: `ollama run llama3`

3. **서버 실행**
   ```bash
   uvicorn app:app --reload
   ```

4. **API 사용 예시**
   - `/upload_chat` 엔드포인트에 `file`(엑셀)과 `question`(질문) 폼 데이터로 POST 요청

---

## 📝 API 예시

- **POST** `/upload_chat`
  - **file**: 엑셀 파일(.xlsx)
  - **question**: 자연어 질문(예: "2023년 매출이 가장 높은 달은?")
  - **응답**: `{ "answer": "..." }`

---

## 📂 폴더 구조

```
ExcelInsightAI/
├── app.py           # FastAPI 서버 및 핵심 로직
├── prompt.txt       # LLM 프롬프트 템플릿
├── README.md        # 프로젝트 설명서
└── ...              # 기타 파일
```

---

## ⚠️ 참고 사항

- **JSONCONVERT.BAS** 파일은 본 프로젝트의 직접 제작 파일이 아니며, 외부 오픈소스인 [VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON)에서 제공된 파일임을 명확히 밝힙니다. 해당 프로젝트는 MIT 라이선스로 배포되고 있습니다.
- 본 프로젝트의 핵심 코드는 생성형 AI로 작성하였습니다.

---

## 💡 커스터마이징 팁

- `prompt.txt` 파일을 수정해 원하는 답변 스타일로 LLM을 튜닝할 수 있습니다.
- Ollama 모델을 변경하려면 `app.py`의 `model="llama3"` 부분을 원하는 모델명으로 바꿔주세요.

---

## 🖥️ 개발/테스트 환경

- Python 3.8+
- FastAPI
- LangChain
- Ollama (로컬 LLM)
- FAISS

---

## 🙏 기여 및 문의

- 이 프로젝트는 자유롭게 포크/수정 가능합니다.
- 버그 제보, 기능 요청, 문의는 이슈로 남겨주세요!

---

**엑셀 데이터, 이제 AI에게 물어보세요!** 
