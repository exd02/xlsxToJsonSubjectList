# Extrator JSON de Alunos em Recuperação

Este repositório fornece duas ferramentas em Python para extração e exportação de notas de planilhas Excel (.xlsx) no formato JSON:

- **extract_operations.py**: biblioteca de funções para leitura de planilhas, normalização de cabeçalhos e geração de dicionários/JSON com as disciplinas reprovadas por aluno.  
- **main.py**: aplicação gráfica web/desktop usando Flet que permite ao usuário selecionar arquivos Excel, configurar colunas e linhas, validar os dados inseridos e exportar o JSON final.

--- 

## Informações de entrada e saída
<p align="center">
  <img src="https://github.com/user-attachments/assets/7be2b0c2-8040-428c-a4de-75d5b2008315" />
</p>

O programa recebe como entrada:
- a localização das planilhas ex.: `C:/dados/planilha.xlsx`
- colunas aonde estão as notas dos alunos (na imagem de exemplo, `F,J,N,R,V,...,BR`)
- a linha aonde está o primeiro aluno (na imagem de exemplo, 3)
- a linha aonde está o último aluno (na imagem de exemplo, 6)

O sistema exporta para um arquivo `AlunosEmRecuperacao.json`, no estilo abaixo, reunindo as disciplinas que cada aluno teve media < 6:
```json
{
    "Agropecuaria": {
        "0": [
            "quimica_i",
            "fisica_i",
            "ingles_i"
        ],
        "1": [
            "quimica_i"
        ],
        "2": [
            "quimica_i",
            "fisica_i"
        ]
    }
    "Informatica": {
        "0": [
            "fisica_i",
            "introducao_a_algoritmos",
            "matematica_i",
            "historia_i",
            "geografia_i",
            "quimica_i"
        ],
        "1": [
            "fisica_i"
        ],
        "2": [
            "historia_i"
        ],
        "3": [
            "fisica_i",
            "banco_de_dados",
            "quimica_i",
            "sociologia_i"
        ]
    }
}
```

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

<p align="center">
  <img src="https://github.com/user-attachments/assets/3bf244b3-cbc8-4c3c-8549-c9f3403c78f1" />
</p>

<p align="center">
  <img src="https://github.com/user-attachments/assets/dff6d1e2-a71a-4896-9065-67acc175dada" />
</p>
