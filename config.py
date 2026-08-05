from pathlib import Path

PASTA_UPLOADS = Path("uploads")
PASTA_UPLOADS.mkdir(exist_ok=True)

ARQUIVO_STATUS = Path("documentos_desativados.json")
ARQUIVO_INDICE = Path("indice_documentos.json")

FORMATOS_ACEITOS = {".txt", ".docx", ".pdf"}

TAMANHO_TRECHO = 1200
SOBREPOSICAO = 200
QUANTIDADE_TRECHOS = 6