import html
import os
import shutil
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse
from openai import OpenAI

from config import FORMATOS_ACEITOS, PASTA_UPLOADS
from documentos import (
    carregar_desativados,
    documentos_ativos,
    extrair_texto,
    listar_documentos,
    salvar_desativados,
)
from indice import (
    buscar_trechos_relevantes,
    recriar_indice,
    referencia_do_trecho,
    remover_do_indice,
)

load_dotenv(".env.local")

app = FastAPI()


ESTILOS_RESPOSTA = {
    "clara": """
Use uma linguagem clara, acolhedora e conversacional.
Explique os termos mais difíceis de maneira simples.
""",
    "teologica": """
Use uma linguagem teológica, com atenção a doutrina, fé e interpretação bíblica.
Defina termos teológicos somente quando houver base nos documentos.
""",
    "academica": """
Use uma linguagem acadêmica, organizada e analítica.
Apresente distinções, limites e nuances com precisão.
""",
}


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
        <form action="/reindexar" method="post">
            <button type="submit">Preparar biblioteca para consulta</button>
        </form>
    </div>
    """


def referencias_html(trechos):
    referencias = {}

    for trecho in trechos:
        nome = trecho["arquivo"]
        pagina = trecho.get("pagina")

        if nome not in referencias:
            referencias[nome] = set()

        if pagina:
            referencias[nome].add(pagina)

    itens = []

    for nome in sorted(referencias):
        paginas = sorted(referencias[nome])

        if paginas:
            detalhe = "Páginas consultadas: " + ", ".join(
                str(pagina) for pagina in paginas
            )
        else:
            detalhe = "Trechos consultados."

        itens.append(
            f"<li><strong>{html.escape(nome)}</strong> — "
            f"{html.escape(detalhe)}</li>"
        )

    return (
        "<div class='conteudo'><h2>Referências consultadas</h2>"
        f"<ul>{''.join(itens)}</ul></div>"
    )


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
            input, select, button {{
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
                <label for="estilo">Como você prefere a resposta?</label>
                <br>
                <select name="estilo" id="estilo">
                    <option value="clara" selected>Clara e acolhedora</option>
                    <option value="teologica">Teológica</option>
                    <option value="academica">Acadêmica</option>
                </select>
                <br>
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
        return pagina(
            "O arquivo foi enviado, mas não foi possível ler o conteúdo."
        )

    if not texto.strip():
        return pagina(
            "O arquivo foi enviado, mas não foi encontrado texto nele."
        )

    conteudo = (
        "<div class='conteudo'><h2>Conteúdo lido</h2>"
        f"{html.escape(texto[:6000])}</div>"
    )

    return pagina(f"Arquivo recebido: {html.escape(nome)}", conteudo)


@app.post("/reindexar", response_class=HTMLResponse)
async def reindexar_documentos():
    if not os.getenv("OPENAI_API_KEY"):
        return pagina("A chave da OpenAI não foi encontrada na configuração.")

    try:
        quantidade_documentos, quantidade_trechos = recriar_indice()
    except Exception as erro:
        print(f"Erro ao preparar a biblioteca: {erro}")
        return pagina(
            "Não foi possível preparar a biblioteca agora. "
            "Confira o terminal e tente novamente."
        )

    if not quantidade_trechos:
        return pagina(
            "Não encontrei texto nos documentos ativos para preparar."
        )

    return pagina(
        "Biblioteca preparada com "
        f"{quantidade_documentos} documento(s) e "
        f"{quantidade_trechos} trecho(s)."
    )


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
    remover_do_indice(nome)

    return pagina(f"Documento removido: {html.escape(nome)}")


@app.post("/perguntar", response_class=HTMLResponse)
async def perguntar(
    pergunta: str = Form(...),
    estilo: str = Form("clara"),
):
    if not os.getenv("OPENAI_API_KEY"):
        return pagina(
            resposta=(
                "<div class='resposta'>"
                "<h2>Configuração pendente</h2>"
                "A chave da OpenAI não foi encontrada."
                "</div>"
            )
        )

    trechos = buscar_trechos_relevantes(pergunta)

    if not trechos:
        return pagina(
            resposta=(
                "<div class='resposta'>"
                "<h2>Sem documentos para consulta</h2>"
                "Envie documentos ativos e clique em "
                "<strong>Preparar biblioteca para consulta</strong>."
                "</div>"
            )
        )

    contexto = "\n\n".join(
        (
            f"--- REFERÊNCIA: {referencia_do_trecho(trecho)} ---\n"
            f"{trecho['trecho']}"
        )
        for trecho in trechos
    )

    instrucao_estilo = ESTILOS_RESPOSTA.get(
        estilo,
        ESTILOS_RESPOSTA["clara"],
    )

    prompt = f"""
Você é um assistente de consulta de documentos.

Use exclusivamente os trechos recuperados abaixo.
Não use conhecimento externo, suposições ou informações que não estejam
claramente apoiadas pelos documentos.

Para cada afirmação importante, indique a referência fornecida no trecho:
- Em PDFs, cite no formato: (nome do documento, página X).
- Em Word ou texto simples, cite apenas o nome do documento.
- Nunca invente uma página ou uma referência.

Regras de fidelidade:
- Só diga que uma ideia está presente nos dois documentos quando houver apoio claro nos dois.
- Se os documentos abordarem o mesmo assunto por perspectivas ou exemplos diferentes, chame isso de "tema relacionado".
- Se uma informação aparecer em apenas um documento, deixe isso explícito.
- Se não houver base suficiente, responda:
  "Não encontrei base suficiente nos documentos consultados."
- Só faça comparação quando a pergunta pedir comparação.
- Ao responder sobre o pensamento de um autor, use:
  "Com base nos documentos consultados..."

ESTILO DA RESPOSTA:
{instrucao_estilo}

TRECHOS RECUPERADOS:
{contexto}

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
        "<div class='conteudo'><h2>Pergunta</h2>"
        f"{html.escape(pergunta)}</div>"
        "<div class='resposta'><h2>Resposta</h2>"
        f"{html.escape(texto_resposta)}</div>"
        f"{referencias_html(trechos)}"
    )

    return pagina(resposta=resposta)