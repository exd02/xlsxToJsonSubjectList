# Extrator JSON de Alunos em Recuperação

Este repositório fornece duas ferramentas em Python para extração e exportação de notas de planilhas Excel (.xlsx) no formato JSON:

- **extract_operations.py**: biblioteca de funções para leitura de planilhas, normalização de cabeçalhos e geração de dicionários/JSON com as disciplinas reprovadas por aluno.  
- **main.py**: aplicação gráfica web/desktop usando Flet que permite ao usuário selecionar arquivos Excel, configurar colunas e linhas, validar os dados inseridos e exportar o JSON final.

---

## Pré-requisitos

Python (versão 3.13.3)
Bibliotecas Python:  
`pip install flet openpyxl` (versão 3.1.5)

---

## Como executar / preview

1. Clone ou copie este repositório.  
2. Instale as dependências:  
        pip install flet openpyxl  
3. Execute a interface gráfica:  
        python main.py  
4. Na janela que abrir:  
    - Clique em **Upload** e selecione uma ou mais planilhas .xlsx (para selecionar várias planilhas basta ir clicando de uma em uma com o CTRL pressionado).  
    - Clique em **Carregar** para passar à configuração.  
    - Ajuste **Cabeçalhos**, **Linha primeiro Aluno** e **Linha último Aluno** para cada planilha.  
    - Clique em **Carregar dados em JSON**.  
5. O arquivo **AlunosEmRecuperacao.json** será gerado e salvo na raiz do projeto, contendo o mapeamento de cada aluno para as disciplinas reprovadas.

---