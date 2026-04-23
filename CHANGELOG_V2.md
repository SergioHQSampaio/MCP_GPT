# Changelog da integracao v2

## O que mudou nesta iteracao

- `convert_file` foi implementado como primeira camada de conversao
- `search_items` passou a suportar busca recursiva real na subarvore
- a documentacao foi atualizada para refletir o novo estado do conector

## O que ja entra nesta integracao

- leitura semantica para DOCX, XLSX, PPTX e PDF textual
- tratamento de texto simples ampliado
- escrita semantica para DOCX, XLSX e PPTX
- preservacao do formato pedido quando houver rota tecnica implementada
- exportacao de arquivos Google nativos para formatos externos comuns
- conversoes basicas para Google Doc, Google Sheet e Google Slides

## O que continua pendente

- suporte semantico pleno a Google Slides nativo
- OCR para PDF escaneado
- leitura visual rica e transformacao orientada a conteudo para imagens
- ampliacao e endurecimento das conversoes mais complexas


## Etapa seguinte
- leitura semantica basica de Google Slides nativo via API
- escrita de Google Slides com slides, titulo e corpo
- conversao para Google Slides agora materializa conteudo real
