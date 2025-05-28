# 📄 OCR PDF Organizer – Organização e Renomeação de PDFs com OCR

Projeto Python que realiza **OCR (Reconhecimento Óptico de Caracteres)** em documentos PDF para extrair nomes de colaboradores e reorganizar os arquivos automaticamente. Ideal para setores de RH, clínicas ou empresas que recebem grandes volumes de exames e precisam classificá-los por nome de forma automatizada.

## 🔧 Tecnologias Utilizadas

- Python 3.x
- [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/en/latest/)
- [ocrmypdf](https://ocrmypdf.readthedocs.io/en/latest/) (OCR em PDFs)
- Expressões Regulares (Regex)
- `shutil`, `os`, `pathlib` (manipulação de arquivos)

## 📁 Estrutura do Projeto

- `renomear_pdf.py`: script principal com funções OCR, renomeação e análise de texto.
- Listas internas com centenas de padrões para identificação e reordenação automática.
- OCR integrado diretamente ao fluxo, com fallback inteligente.
- Suporte a múltiplos padrões de nome de arquivos PDF usados em ambientes médicos.

## ▶️ Como Executar

Certifique-se de ter o Python 3 e o pacote `ocrmypdf` instalado:
   pip install pymupdf
   sudo apt install ocrmypdf  # Ou use Chocolatey no Windows
