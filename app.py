import io
import os
import pandas as pd
import pypdf
import streamlit as st
from openai import AzureOpenAI

# Configuração inicial da página do Streamlit
st.set_page_config(
    page_title="ChatBot RAG - Microsoft Foundry", page_icon="🤖", layout="centered"
)

# 1. Interface e Estilo (CSS Customizado para Borda e Aparência Profissional)
st.markdown(
    """
    <style>
        .stApp {
            background-color: #f4f6f9;
        }
        [data-testid="stSidebar"] {
            background-color: #ffffff;
            border-right: 1px solid #e2e8f0;
        }
        .main-content-box {
            border: 2px solid #cbd5e1;
            border-radius: 12px;
            padding: 20px;
            background-color: #ffffff;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            margin-bottom: 20px;
        }
        h1 {
            color: #1e293b;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# 2. Configuração da API (Microsoft Foundry / Azure AI)
API_KEY = "6t2pTNsmNafDyo53MDAFIrvvQRBZXP7dLcr5j3TLxMS34LQVq3ZUJQQJ99CHACHYHv6XJ3w3AAAAACOGT7os"
DEPLOYMENT_NAME = "gpt-4.1-mini"
ENDPOINT = "https://marcelomaffeis-05082026-resource.services.ai.azure.com/"

# Inicialização segura do cliente Azure OpenAI
client = None
try:
  client = AzureOpenAI(
      api_key=API_KEY, api_version="2024-02-01", azure_endpoint=ENDPOINT
  )
except Exception as e:
  st.error(f"Erro ao configurar o cliente da API: {e}")

# Título da Aplicação
st.title("🤖 ChatBot Inteligente com RAG")
st.markdown(
    "Faça upload de um documento (**Excel**, **PDF** ou **TXT**) na barra"
    " lateral para utilizá-lo como contexto."
)

# Inicialização do Session State para persistência
if "document_text" not in st.session_state:
  st.session_state.document_text = ""

if "messages" not in st.session_state:
  st.session_state.messages = []

# 3. Upload de Arquivos e Sistema RAG na Barra Lateral
st.sidebar.header("📁 Base de Conhecimento")
uploaded_file = st.sidebar.file_uploader(
    "Envie seu arquivo (Excel, PDF ou TXT)", type=["xlsx", "xls", "pdf", "txt"]
)

# Processamento seguro do arquivo carregado
if uploaded_file is not None:
  file_extension = uploaded_file.name.split(".")[-1].lower()
  try:
    if file_extension == "txt":
      st.session_state.document_text = uploaded_file.getvalue().decode("utf-8")
      st.sidebar.success(f"TXT '{uploaded_file.name}' carregado com sucesso!")

    elif file_extension == "pdf":
      pdf_reader = pypdf.PdfReader(io.BytesIO(uploaded_file.getvalue()))
      extracted_text = ""
      for page in pdf_reader.pages:
        text = page.extract_text()
        if text:
          extracted_text += text + "\n"
      st.session_state.document_text = extracted_text
      st.sidebar.success(f"PDF '{uploaded_file.name}' carregado com sucesso!")

    elif file_extension in ["xlsx", "xls"]:
      # Leitura de planilhas Excel convertendo em texto estruturado para o RAG
      df = pd.read_excel(uploaded_file)
      st.session_state.document_text = df.to_string(index=False)
      st.sidebar.success(f"Excel '{uploaded_file.name}' carregado com sucesso!")

  except Exception as e:
    st.sidebar.error(f"Erro ao processar o arquivo: {e}")
    st.session_state.document_text = ""

# Exibir indicador de contexto ativo na barra lateral
if st.session_state.document_text:
  st.sidebar.info(
      "📄 Contexto carregado e pronto para consultas no chat!"
  )

# 4. Lógica de Chat
for message in st.session_state.messages:
  with st.chat_message(message["role"]):
    st.markdown(message["content"])

# Capturar entrada do usuário
if prompt := st.chat_input(
    "Digite sua mensagem ou pergunta sobre a planilha/documento..."
):
  st.session_state.messages.append({"role": "user", "content": prompt})
  with st.chat_message("user"):
    st.markdown(prompt)

  # Construção do Prompt do Sistema (RAG Contextual)
  system_prompt = (
      "Você é um assistente virtual especialista e prestativo em análise de"
      " dados e engenharia. Responda de maneira clara e objetiva."
  )

  if st.session_state.document_text:
    system_prompt += (
        "\n\nUtilize obrigatoriamente as informações abaixo extraídas do"
        f" documento/planilha enviado pelo usuário para formular sua"
        f" resposta:\n--- INÍCIO DO CONTEXTO ---\n{st.session_state.document_text}\n--- FIM"
        " DO CONTEXTO ---\nSe a informação necessária não estiver presente no"
        " contexto, informe educadamente."
    )

  # Montagem das mensagens para a API
  messages_payload = [{"role": "system", "content": system_prompt}]
  for m in st.session_state.messages:
    messages_payload.append({"role": m["role"], "content": m["content"]})

  # Chamada à API da Microsoft Foundry (Azure OpenAI)
  with st.chat_message("assistant"):
    with st.spinner("Analisando e gerando resposta..."):
      if client is None:
        st.error("Cliente da API não foi inicializado corretamente.")
      else:
        try:
          response = client.chat.completions.create(
              model=DEPLOYMENT_NAME, messages=messages_payload, temperature=0.7
          )
          bot_response = response.choices[0].message.content
          st.markdown(bot_response)
          st.session_state.messages.append(
              {"role": "assistant", "content": bot_response}
          )
        except Exception as e:
          st.error(f"Erro ao se comunicar com a API do Microsoft Foundry: {e}")
