# Google Drive Bridge MCP

Servidor MCP para conectar o ChatGPT a uma pasta-raiz especifica do Google Drive, com operacoes restritas a essa subarvore e com **suporte orientado a capacidades por formato**.

## Definicao de produto

Este projeto deve ser entendido como uma camada de operacao de arquivos sobre uma subarvore do Google Drive.

Ele atua em tres niveis:

1. **Operacao universal de arquivo**
   - listar
   - localizar
   - obter metadados
   - criar
   - substituir
   - mover
   - renomear
   - lixeira
   - restauracao

2. **Leitura semantica**
   - extrair texto, estrutura e metadados de formatos suportados

3. **Escrita semantica / substituicao / conversao**
   - editar com backend especializado quando houver rota robusta
   - substituir binariamente quando isso for mais seguro
   - converter quando houver rota tecnica confiavel

## O que a versao atual cobre

### Estrutural
- `gdrive_health`
- `list_items`
- `search_items`
- `create_folder`
- `get_file_metadata`
- `move_item`
- `rename_item`
- `trash_item`
- `restore_item`

### Compatibilidade MVP
- `create_doc`
- `read_doc`
- `write_doc`
- `create_sheet`
- `read_sheet_range`
- `write_sheet_range`
- `upload_text_file`
- `upload_base64_file`
- `download_file_metadata`
- `read_text_file`

### Camada unificada nova
- `create_file`
- `read_file`
- `write_file`
- `convert_file`

## Formatos com leitura implementada nesta fase

- Google Docs
- Google Sheets (sumario + preview)
- TXT / MD / JSON / CSV / HTML / XML / YAML
- DOCX
- XLSX
- PPTX
- PDF textual
- imagens com metadados basicos

## Formatos com escrita implementada nesta fase

- Google Docs
- Google Sheets por `range + values`
- texto simples
- DOCX
- XLSX
- PPTX
- binarios por substituicao

## Conversoes implementadas nesta fase

### Exportacao de arquivos Google nativos
- Google Docs -> DOCX / PDF / TXT / MD
- Google Sheets -> XLSX / PDF / CSV
- Google Slides -> PPTX / PDF

### Conversoes internas basicas
- TXT / MD / DOCX / PDF textual -> Google Doc
- XLSX / Google Sheet / texto tabular simples -> Google Sheet
- PPTX / TXT / DOCX / PDF textual -> Google Slides estrutural basico
- formatos nao Google -> outro formato serializavel quando houver parser + serializer implementados

## Melhorias tecnicas novas

- `search_items` agora suporta busca recursiva real em subarvore
- `convert_file` foi introduzido como camada de conversao
- a arquitetura continua preservando a versao legada para compatibilidade

## Limites atuais

- Google Slides nativo ainda nao tem leitura/escrita semantica completa nesta fase; a criacao estrutural esta apenas basica
- PDF esta tratado para extracao textual; OCR ainda nao entrou
- imagens ainda estao em nivel de metadados, nao de analise visual rica
- algumas conversoes sao deliberadamente por reconstrucao, nao por edicao interna do formato

## Regra central de seguranca

O servidor **nunca** deve operar fora de `GDRIVE_ROOT_FOLDER_ID`.

Toda operacao por `file_id` passa por validacao de ancestralidade.
Todo caminho informado pelo ChatGPT e resolvido relativamente a essa pasta-raiz.

## Requisitos

- Python 3.11+
- Google Drive API
- Google Docs API
- Google Sheets API
- OAuth de usuario ou service account
- endpoint HTTPS remoto para uso real no ChatGPT

## Dependencias Python

Veja `requirements.txt`.

## Execucao local

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python server_mcp.py
```

Por padrao:

```text
http://0.0.0.0:8787/mcp
```


## Atualizacao recente

A versao atual ja faz leitura semantica basica de Google Slides nativo e escrita estrutural simples (titulo + corpo) via API do Slides.
