import json
import math

from openai import OpenAI

from config import (
    ARQUIVO_INDICE,
    PASTA_UPLOADS,
    QUANTIDADE_TRECHOS,
    SOBREPOSICAO,
    TAMANHO_TRECHO,
)
from documentos import (
    carregar_desativados,
    documentos_ativos,
    extrair_partes_referenciadas,
)


def carregar_indice():
    if not ARQUIVO_INDICE.exists():
        return []

    try:
        return json.loads(ARQUIVO_INDICE.read_text(encoding="utf-8"))
    except Exception:
        return []


def salvar_indice(indice):
    ARQUIVO_INDICE.write_text(
        json.dumps(indice, ensure_ascii=False),
        encoding="utf-8",
    )


def dividir_em_trechos(texto):
    texto = " ".join(texto.split())

    if not texto:
        return []

    trechos = []
    passo = TAMANHO_TRECHO - SOBREPOSICAO

    for inicio in range(0, len(texto), passo):
        trecho = texto[inicio : inicio + TAMANHO_TRECHO]

        if trecho:
            trechos.append(trecho)

        if inicio + TAMANHO_TRECHO >= len(texto):
            break

    return trechos


def gerar_embeddings(textos):
    cliente = OpenAI()
    embeddings = []

    for inicio in range(0, len(textos), 100):
        lote = textos[inicio : inicio + 100]
        resultado = cliente.embeddings.create(
            model="text-embedding-3-small",
            input=lote,
        )
        embeddings.extend(item.embedding for item in resultado.data)

    return embeddings


def recriar_indice():
    documentos = documentos_ativos()
    registros = []

    for documento in documentos:
        partes = extrair_partes_referenciadas(documento)

        for parte in partes:
            for trecho in dividir_em_trechos(parte["texto"]):
                registros.append(
                    {
                        "arquivo": documento.name,
                        "pagina": parte["pagina"],
                        "trecho": trecho,
                    }
                )

    if not registros:
        salvar_indice([])
        return 0, 0

    embeddings = gerar_embeddings(
        [registro["trecho"] for registro in registros]
    )

    for registro, embedding in zip(registros, embeddings):
        registro["embedding"] = embedding

    salvar_indice(registros)
    return len(documentos), len(registros)


def remover_do_indice(nome):
    indice = carregar_indice()
    indice = [
        registro
        for registro in indice
        if registro.get("arquivo") != nome
    ]
    salvar_indice(indice)


def similaridade(vetor_a, vetor_b):
    produto = sum(a * b for a, b in zip(vetor_a, vetor_b))
    norma_a = math.sqrt(sum(a * a for a in vetor_a))
    norma_b = math.sqrt(sum(b * b for b in vetor_b))

    if not norma_a or not norma_b:
        return 0

    return produto / (norma_a * norma_b)


def buscar_trechos_relevantes(pergunta):
    desativados = carregar_desativados()
    indice = carregar_indice()

    registros_ativos = [
        registro
        for registro in indice
        if registro.get("arquivo") not in desativados
        and (PASTA_UPLOADS / registro.get("arquivo", "")).exists()
    ]

    if not registros_ativos:
        return []

    vetor_pergunta = gerar_embeddings([pergunta])[0]

    for registro in registros_ativos:
        registro["pontuacao"] = similaridade(
            vetor_pergunta,
            registro["embedding"],
        )

    return sorted(
        registros_ativos,
        key=lambda registro: registro["pontuacao"],
        reverse=True,
    )[:QUANTIDADE_TRECHOS]


def referencia_do_trecho(trecho):
    nome = trecho["arquivo"]
    pagina = trecho.get("pagina")

    if pagina:
        return f"{nome}, página {pagina}"

    return nome