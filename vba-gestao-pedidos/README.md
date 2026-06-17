# Projeto em Visual Basic para Gestão de Pedidos de Compra

## Objetivo 
### A automação em VBA para integração e tratamento de dados de compras entre Pipefy (CSV) e Excel.
- A Automação padroniza por áreas no modelo da empresa.
- Permanência da informação com um software essencial.
- Banco de dados.

## Problema do processo
 O processo de atualização da planilha de compras era manual, exigindo a leitura linha a linha de relatórios exportados do Pipefy. Isso gerava:
- Alto tempo gasto em tarefas repetitivas.
- Risco de erros de digitação e duplicação de dados.
- Dificuldade em padronizar nomes de departamentos e status.
- Computador da empresa limita utilização de softwares/IDE's.
- Impossibilidade de migração total ao Pipefy por conta de processo interno.

### Problema institucional
- A escolha da aplicação é feita pela resistência a novas tecnologias em empresa tradicional.
- A variação orçamentária da paraestatal gera instabilidade nos softwares disponibilizados pelas equipes, sendo o Excel uma forma muito simples e segura de controlar os dados.
- Os controles geralmente são individuais, facilitando a utilização de macros, não há outros colaboradores envolvidos na gestão do processo. 

## Solução
Desenvolvi um script em **VBA (Visual Basic for Applications)** que atua como um ETL (Extract, Transform, Load) leve dentro do próprio Excel. O sistema lê o CSV bruto, limpa os dados e atualiza a base principal automaticamente.

### Funções principais
**Importação Inteligente:** Lê arquivos CSV (UTF-8) exportados do Pipefy.
**Lógica "Upsert":** Verifica se o pedido já existe. Se existir, atualiza o status; se não, cria uma nova linha.
**Sanitização de Dados:** Padroniza cabeçalhos, remove acentos e corrige variações de texto (ex: "Nº" vs "N°").
**Gestão Visual:** Novas requisições são destacadas automaticamente em **azul** para fácil identificação.
**Prevenção de Erros:** Protege dados sensíveis (ex: não sobrescreve o solicitante original se o sistema tiver sido atualizado por um "bot").

## Ferramentas aplicadas
**Linguagem:** VBA (Excel).
**Manipulação de Dados:** Tratamento de Strings, Arrays e Dictionaries.
**Lógica de Negócio:** Mapeamento de chaves primárias (IDs) para evitar duplicatas.
**UX (Experiência do Usuário):** Automação acionada por botão simples, sem necessidade de conhecimento técnico dos colaboradores.

## Demonstração
1. Código no Visual Basic:
<img width="1763" height="706" alt="{6A76E3E3-7E2D-475A-ACA6-1AF67BEA9AF3}" src="https://github.com/user-attachments/assets/82a7e3d7-5dc0-45fe-9a9b-41a00b491a44" />

2. Execução do código:
<img width="1159" height="434" alt="{7A69154E-9826-4F37-AC13-C9FD89CA987C}" src="https://github.com/user-attachments/assets/9c78fb9c-1000-4cdc-95cc-64f8b39635ae" />
<img width="839" height="508" alt="{C991EDDF-8B6D-4971-B28E-E5645C60A40C}" src="https://github.com/user-attachments/assets/b5d94b82-93f7-4127-a4dd-a3d6e50c9261" />

3. Conclusão:
<img width="856" height="361" alt="{C905FAFB-6483-4010-B037-40E5FD97B1C6}" src="https://github.com/user-attachments/assets/790e7b60-10a2-449f-8800-8903ac15b1f9" />
<img width="1658" height="286" alt="{6B71E9E5-AB63-4FEE-8711-72F778B0DC44}" src="https://github.com/user-attachments/assets/c7fea7c1-bc1f-4bc5-b907-2164f1fdf698" />


## Como Usar
1. Exporte o relatório do Pipefy em `.csv`.
2. Abra a planilha deste repositório (`.xlsm`).
3. Clique no botão **"Atualizar Dados"** (ou execute a macro `AtualizarComprasComPipefy`).
4. Selecione o arquivo CSV.
5. Pronto! O sistema informará quantas linhas foram criadas ou atualizadas. :)

--- 
Projeto desenvolvido em período de estágio. 

At. 16/06/2026
