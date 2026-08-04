from datetime import datetime
from pathlib import Path
import html
import json
import os
import shutil

from docx import Document
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse
from openai import OpenAI
from pypdf import PdfReader

load_dotenv(".env.local")

app = FastAPI()

PASTA_UPLOADS = Path("uploads")
PASTA_UPLOADS.mkdir(exist_ok=True)

ARQUIVO_STATUS = Path("documentos_desativados.json")
FORMATOS_ACEITOS = {".txt", ".docx", ".pdf"}


def carregar_desativados():
    if not ARQUIVO_STATUS.exists():
        return set()

    try:
        return set(json.loads(ARQUIVO_STATUS.read_text(encoding="utf-8")))
    except Exception:
        return set()


def salvar_desativados(desativados):
    ARQUIVO_STATUS.write_text(
        json.dumps(sorted(desativados), ensure_ascii=False),
        encoding="utf-8",
    )


def listar_documentos():
    desativados = carregar_desativados()
    documentos = []

    for arquivo in PASTA_UPLOADS.iterdir():
        if arquivo.is_file() and arquivo.suffix.lower() in FORMATOS_ACEITOS:
            documentos.append(
                {
                    "nome": arquivo.name,
                    "tipo": arquivo.suffix.upper().replace(".", ""),
                    "data": datetime.fromtimestamp(
                        arquivo.stat().st_mtime
                    ).strftime("%d/%m/%Y %H:%M"),
                    "ativo": arquivo.name not in desativados,
                }
            )

    return sorted(documentos, key=lambda documento: documento["data"], reverse=True)


def documentos_html():
    documentos = listar_documentos()

    if not documentos:
        return """
        <div class="documentos">
            <h2>Documentos enviados</h2>
            <p>Nenhum documento enviado ainda.</p>
        </div>
        """

    itens = []

    for documento in documentos:
        nome = html.escape(documento["nome"])
        estado = "Ativo" if documento["ativo"] else "Desativado"
        acao = "Desativar" if documento["ativo"] else "Reativar"

        itens.append(
            f"""
            <div class="documento">
                <div>
                    <strong>{nome}</strong><br>
                    <span>{documento["tipo"]} · {documento["data"]} · {estado}</span>
                </div>
                <div>
                    <form action="/alternar" method="post" class="acao">
                        <input type="hidden" name="nome" value="{nome}">
                        <button type="submit" class="secundario">{acao}</button>
                    </form>
                    <form action="/remover" method="post" class="acao"
                          onsubmit="return confirm('Remover este documento?')">
                        <input type="hidden" name="nome" value="{nome}">
                        <button type="submit" class="remover">Remover</button>
                    </form>
                </div>
            </div>
            """
        )

    return f"""
    <div class="documentos">
        <h2>Documentos enviados</h2>
        {''.join(itens)}
    </div>
    """


def pagina(mensagem="", conteudo="", resposta=""):
    return f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <title>Assistente de Documentos</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                max-width: 760px;
                margin: 60px auto;
                padding: 0 20px;
                color: #1f2937;
            }}
            h1 {{ color: #2563eb; text-align: center; }}
            h2 {{ margin-top: 0; }}
            .caixa, .documentos {{
                margin-top: 28px;
                padding: 28px;
                border-radius: 12px;
                background: #eff6ff;
            }}
            .caixa {{
                border: 2px dashed #93c5fd;
                text-align: center;
            }}
            input, button {{
                margin: 10px;
                padding: 12px;
                font-size: 16px;
            }}
            input[type="text"] {{ width: 70%; }}
            button {{
                border: 0;
                border-radius: 8px;
                background: #2563eb;
                color: white;
                cursor: pointer;
            }}
            .secundario {{ background: #64748b; }}
            .remover {{ background: #dc2626; }}
            .mensagem {{ color: #166534; font-weight: bold; }}
            .conteudo, .resposta {{
                margin-top: 28px;
                padding: 20px;
                border-radius: 12px;
                background: #f8fafc;
                white-space: pre-wrap;
            }}
            .resposta {{ border-left: 5px solid #2563eb; }}
            .documento {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                gap: 16px;
                padding: 16px 0;
                border-top: 1px solid #cbd5e1;
            }}
            .documento:first-of-type {{ border-top: 0; }}
            .documento span {{ color: #475569; font-size: 14px; }}
            .acao {{ display: inline; }}
            .acao button {{ margin: 4px; padding: 8px 12px; font-size: 14px; }}
        </style>
    </head>
    <body>
        <h1>Assistente de Documentos</h1>

        <div class="caixa">
            <h2>Enviar documento</h2>
            <p>Formatos aceitos: texto, Word e PDF.</p>
            <form action="/enviar" method="post" enctype="multipart/form-data">
                <input type="file" name="arquivo"
                       accept=".txt,.docx,.pdf" required>
                <br>
                <button type="submit">Enviar documento</button>
            </form>
            <p class="mensagem">{mensagem}</p>
        </div>

        {documentos_html()}

        <div class="caixa">
            <h2>Faça uma pergunta</h2>
            <form action="/perguntar" method="post">
                <input type="text" name="pergunta"
                       placeholder="Ex.: Faça um resumo do documento"
                       required>
                <br>
                <button type="submit">Perguntar</button>
            </form>
        </div>

        {conteudo}
        {resposta}
    </body>
    </html>
    """


def extrair_texto(caminho):
    extensao = caminho.suffix.lower()

    if extensao == ".txt":
        return caminho.read_text(encoding="utf-8", errors="replace")

    if extensao == ".docx":
        documento = Document(str(caminho))
        partes = [paragrafo.text for paragrafo in documento.paragraphs]

        for tabela in documento.tables:
            for linha in tabela.rows:
                partes.append(" | ".join(celula.text for celula in linha.cells))

        return "\n".join(partes)

    if extensao == ".pdf":
        leitor = PdfReader(str(caminho))
        return "\n".join(pagina.extract_text() or "" for pagina in leitor.pages)

    raise ValueError("Formato de arquivo não aceito.")


def documentos_ativos():
    desativados = carregar_desativados()

    return sorted(
        [
            arquivo
            for arquivo in PASTA_UPLOADS.iterdir()
            if (
                arquivo.is_file()
                and arquivo.suffix.lower() in FORMATOS_ACEITOS
                and arquivo.name not in desativados
            )
        ],
        key=lambda arquivo: arquivo.stat().st_mtime,
    )


@app.get("/", response_class=HTMLResponse)
def inicio():
    return pagina()


@app.post("/enviar", response_class=HTMLResponse)
async def enviar_documento(arquivo: UploadFile = File(...)):
    nome = Path(arquivo.filename or "arquivo").name
    destino = PASTA_UPLOADS / nome

    if destino.suffix.lower() not in FORMATOS_ACEITOS:
        return pagina("Formato não aceito. Envie texto, Word ou PDF.")

    with destino.open("wb") as arquivo_destino:
        shutil.copyfileobj(arquivo.file, arquivo_destino)

    desativados = carregar_desativados()
    desativados.discard(nome)
    salvar_desativados(desativados)

    try:
        texto = extrair_texto(destino)
    except Exception as erro:
        print(f"Erro ao ler o documento: {erro}")
        return pagina("O arquivo foi enviado, mas não foi possível ler o conteúdo.")

    if not texto.strip():
        return pagina("O arquivo foi enviado, mas não foi encontrado texto nele.")

    conteudo = (
        "<div class='conteudo'><h2>Conteúdo lido</h2>"
        f"{html.escape(texto[:6000])}</div>"
    )

    return pagina(f"Arquivo recebido: {html.escape(nome)}", conteudo)


@app.post("/alternar", response_class=HTMLResponse)
async def alternar_documento(nome: str = Form(...)):
    nome = Path(nome).name
    arquivo = PASTA_UPLOADS / nome

    if not arquivo.exists():
        return pagina("Documento não encontrado.")

    desativados = carregar_desativados()

    if nome in desativados:
        desativados.remove(nome)
        mensagem = f"Documento reativado: {html.escape(nome)}"
    else:
        desativados.add(nome)
        mensagem = f"Documento desativado: {html.escape(nome)}"

    salvar_desativados(desativados)
    return pagina(mensagem)


@app.post("/remover", response_class=HTMLResponse)
async def remover_documento(nome: str = Form(...)):
    nome = Path(nome).name
    arquivo = PASTA_UPLOADS / nome

    if arquivo.exists():
        arquivo.unlink()

    desativados = carregar_desativados()
    desativados.discard(nome)
    salvar_desativados(desativados)

    return pagina(f"Documento removido: {html.escape(nome)}")


@app.post("/perguntar", response_class=HTMLResponse)
async def perguntar(pergunta: str = Form(...)):
    documentos = documentos_ativos()

    if not documentos:
        return pagina(
            resposta=(
                "<div class='resposta'>"
                "Envie ou reative um documento para eu consultar."
                "</div>"
            )
        )

    if not os.getenv("OPENAI_API_KEY"):
        return pagina(
            resposta=(
                "<div class='resposta'>"
                "A chave da OpenAI não foi encontrada na configuração."
                "</div>"
            )
        )

    partes = []

    for documento in documentos:
        try:
            texto = extrair_texto(documento)
        except Exception as erro:
            print(f"Erro ao ler o documento {documento.name}: {erro}")
            continue

        if texto.strip():
            partes.append(
                f"\n--- DOCUMENTO: {documento.name} ---\n"
                f"{texto[:8000]}"
            )

    if not partes:
        return pagina(
            resposta=(
                "<div class='resposta'>"
                "Não foi possível ler os documentos ativos."
                "</div>"
            )
        )

    contexto = "\n".join(partes)

    prompt = f"""
Você é um assistente de consulta de documentos.
Considere todos os documentos abaixo para responder.
Quando comparar informações, informe de quais documentos elas vieram.
Se a resposta não estiver nos documentos, diga claramente que ela não foi encontrada.

DOCUMENTOS:
{contexto[:24000]}

PERGUNTA:
{pergunta}
"""

    try:
        cliente = OpenAI()
        resultado = cliente.responses.create(
            model="gpt-5.6-luna",
            input=prompt,
        )
        texto_resposta = resultado.output_text
    except Exception as erro:
        print(f"Erro ao consultar a OpenAI: {erro}")
        texto_resposta = (
            "Não foi possível consultar a IA agora. "
            "Confira a conexão e tente novamente."
        )

    resposta = (
        "<div class='resposta'><h2>Resposta</h2>"
        f"{html.escape(texto_resposta)}</div>"
    )

    return pagina(resposta=resposta)