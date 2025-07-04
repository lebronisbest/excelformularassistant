from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import JSONResponse
from langchain_ollama import OllamaEmbeddings, ChatOllama
from langchain_community.vectorstores import FAISS
from langchain.prompts import PromptTemplate
from langchain_core.documents import Document
from langchain_core.prompts import PromptTemplate as CorePromptTemplate
from langchain_core.runnables import RunnableSequence
import tempfile
import os
import pandas as pd

app = FastAPI()

# 프롬프트 파일 경로
PROMPT_PATH = os.path.join(os.path.dirname(__file__), 'prompt.txt')

def load_prompt(path):
    with open(path, encoding='utf-8') as f:
        return f.read()

# 엑셀을 마크다운 문서로 변환
def excel_to_context(path):
    df = pd.read_excel(path)
    header = df.columns.tolist()
    col_letters = [chr(65 + i) for i in range(len(header))]
    new_header = [f"{col}:{name}" for col, name in zip(col_letters, header)]
    header_line = "| " + " | ".join(new_header) + " |"
    sep_line = "| " + " | ".join(["---"] * len(new_header)) + " |"
    data_lines = ["| " + " | ".join(map(str, row)) + " |" for row in df.itertuples(index=False)]
    return "\n".join([header_line, sep_line] + data_lines)

@app.post("/upload_chat")
async def upload_chat(file: UploadFile, question: str = Form(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name

    answer = ""
    try:
        print("[INFO] 엑셀 파일 업로드: ", file.filename)
        context = excel_to_context(tmp_path)
        print("[INFO] LangChain 문서(context) 변환 결과:")
        print(context)

        # 벡터 DB 준비 (필요시 유지, 아니면 삭제 가능)
        docs = [Document(page_content=context)]
        db = FAISS.from_documents(docs, OllamaEmbeddings(model="nomic-embed-text"))

        # 한 번만 LLM 호출 (answer 프롬프트)
        answer_prompt = CorePromptTemplate(
            template=load_prompt(PROMPT_PATH),
            input_variables=["context", "question"]
        )
        print("[INFO] 질문: ", question)
        answer_chain = answer_prompt | ChatOllama(model="llama3")
        answer = answer_chain.invoke({
            "context": context,
            "question": question
        }).content.strip().split("\n")[0]
        print("[INFO] LLM answer(함수): ", answer)

    except Exception as e:
        answer = f"오류: {str(e)}"
        print("[ERROR] ", answer)

    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    return JSONResponse(content={"answer": answer})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8000)