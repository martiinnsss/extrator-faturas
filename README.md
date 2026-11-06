# Extrator de Faturas PDF → Excel

Ferramenta em Python que lê uma pasta de faturas em PDF e junta tudo numa folha de Excel organizada — fornecedor, NIF, número, data, base, IVA e total — com somas automáticas.

**Problema que resolve:** pequenos negócios passam horas por mês a copiar dados de faturas à mão para o Excel ou para enviar ao contabilista. Isto faz o mesmo em segundos e sem erros de cópia.

## Funcionalidades

- Lê todos os PDFs de uma pasta de uma vez
- Extrai fornecedor, NIF, n.º de fatura, data, base tributável, IVA e total
- Entende o formato português de números (`1.234,56 €`) e datas (`dd/mm/aaaa`)
- Gera um Excel formatado: cabeçalho, filtros, valores em euros, ordenação por data e linha de totais com fórmulas
- Faturas que não consegue ler vão para uma folha **"Por rever"** em vez de parar o processo

## Como usar

```bash
pip install -r requirements.txt

# (opcional) gerar faturas de exemplo para testar
python gerar_exemplos.py

# extrair
python extrator.py                        # lê ./faturas → faturas.xlsx
python extrator.py pasta_faturas saida.xlsx
```

## Limitações e próximos passos

- PDFs digitalizados (imagem) precisam de OCR — previsto com `pytesseract`
- Os padrões estão afinados para faturas portuguesas comuns; formatos diferentes podem precisar de ajustes em `PADROES`
- Ideias: interface gráfica simples, envio automático por email ao contabilista, leitura direta da caixa de email

## Tecnologias

Python · pdfplumber · openpyxl · reportlab

---
Feito por **Gonçalo Martins** — estudante de Engenharia Informática e Aplicações.
Precisa de automatizar uma tarefa repetitiva no seu negócio? Fale comigo.
