import streamlit as st 
import os 
import langchain
from langchain_google_genai import ChatGoogleGenerativeAI
from dotenv import load_dotenv
import subprocess
import sys

load_dotenv()
os.environ["GEMINI_API_KEY"] = os.getenv("gem")

model = ChatGoogleGenerativeAI(model="gemini-2.5-flash")

st.title("AI_POWERED_PPT_GENERATOR")
inp = st.text_area("Enter your prompt here:")
prompt = [("system", """You are an AI that outputs only Python code that uses the python-pptx library to create a .pptx presentation.
When the user gives instructions, generate only the final Python code needed to build the PowerPoint — no explanations, no notes, no reasoning, no markdown, no comments, no text before or after the code.
Rules:
Output only executable Python code.
Use python-pptx library only.
Create slides, titles, bullets, images, tables, or anything else exactly as the user requests.
Always save the PowerPoint file at the end using presentation.save("output.pptx") (or another filename if the user specifies).
Do NOT save a PDF version of the file — only .pptx should be created.
Do not include any natural language, explanations, chain-of-thought, markdown, or text outside the code.
If the user input is incomplete, make reasonable assumptions and still output valid code.""")]
prompt.append(("user", inp))

if st.button("Generate PPT"):

    model_response = model.invoke(prompt)

    with open("app.py", "w") as f:
        f.write(model_response.content.strip("```python")) 
    subprocess.run([sys.executable, "app.py"])    

    folder_path = "C:\\Users\\DELL\\Desktop\\PPT_PROJECT\\ppt\\project"
    pptx_files = [f for f in os.listdir(folder_path) if f.endswith(".pptx")]
    if not pptx_files:
        st.error("No PPTX files found!")
    else:
        pptx_files_full = [os.path.join(folder_path, f) for f in pptx_files]
        latest_file = max(pptx_files_full, key=os.path.getctime)

    with open(latest_file, "rb") as f:
            file_bytes = f.read()

    if st.download_button(
        label="Download PPTX",
        data=file_bytes,
        file_name=os.path.basename(latest_file),
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    ):
            os.remove(folder_path + "\\" + os.path.basename(latest_file))