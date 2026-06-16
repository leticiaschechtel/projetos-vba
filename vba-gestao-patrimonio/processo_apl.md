# Aplicação em Processos Patrimôniais

## Objetivo

- Aumentar a capacidade de coleta de informações do formulário do pipefy e transformar os dados coletados em um relatório do processo. 
- Reestruturar a planilha patrimonial para rastreabilidade e manutenção da informação interna.
- Permanência dos dados no Excel e resistência da informação a mudanças institucionais.
  
---
## Estrutura do Pipefy
### Fluxograma do processo desenvolvido no pipefy
<img width="1646" height="512" alt="image" src="https://github.com/user-attachments/assets/4b8b1788-4706-4ffc-95fd-9f7a443e0faf" />

---
## Estrutura do Relatório
### Reestruturação da planilha de patrimonial para receber o relatório.

<img width="1195" height="626" alt="image" src="https://github.com/user-attachments/assets/b6f18d28-6c5f-49a6-bcb9-658ab63496d3" />

- Eram 93 linhas, sem histórico de transferência e sem normalização dos dados.
- A estrutura do Pipefy foi formatada para se tornar um relatório.
- A tabela da planilha foi preparada para receber o relatório do Pipefy e se tornar a base de dados “legado” para consultas gerais.
- Mantém a rastreabilidade futura e histórico de TODAS as transferências informadas e realizadas pelo setor.

### Visualização das atualizações para consulta geral.
<img width="1182" height="486" alt="image" src="https://github.com/user-attachments/assets/a93e3718-f785-49be-b5fb-ab7898ca8b47" />

- Temos 2 abas, uma para visualização e outra base de dados para consultas.
- A aba inventário tem uma linha por movimentação, higienizado e sem repetição.
- A aba histórico tem uma linha por movimentação, nunca apaga, nunca sobrescreve informações relatadas.

---

# CONSTRUÇÃO DAS MACROS NO VBA
##  A lógica para a atualização inteligente.

### MACRO 1 - ATUALIZAÇÃO DE RELATÓRIOS COM DADOS EXTERNOS

<img width="1454" height="222" alt="image" src="https://github.com/user-attachments/assets/a2dd2f61-e720-4ce8-8a00-e2dc2af6d354" />

### MACRO 2 - ATUALIZAÇÃO DE DADOS DIRETAMENTE NA TABELA
<img width="1296" height="222" alt="image" src="https://github.com/user-attachments/assets/96a3a1df-096d-43d6-8a4b-61a01804ad3f" />

---
## MACRO 1 - ATUALIZAÇÃO DE RELATÓRIOS COM DADOS EXTERNOS
### 1° ETAPA - Extração dos relatórios do Pipefy
<img width="1683" height="480" alt="image" src="https://github.com/user-attachments/assets/20a6e503-dd89-4183-9a89-1c528f756fbf" />

### 2° ETAPA: APLICAÇÃO DA MACRO NO EXCEL
<img width="1753" height="527" alt="image" src="https://github.com/user-attachments/assets/6134ab2b-5bcb-40ef-b2cf-c66aac580c34" />

### 3° ETAPA: IMPORTAÇÃO DOS DADOS PELA MACRO
<img width="1378" height="491" alt="image" src="https://github.com/user-attachments/assets/eba3e0ff-9ac2-48ba-8c60-4fa6f601fe85" />

### 4° ETAPA: ATUALIZAÇÃO NA VISUALIZAÇÃO
<img width="1777" height="417" alt="image" src="https://github.com/user-attachments/assets/ab9d7366-2924-4086-8684-67831e624755" />

---
## MACRO 2 - ATUALIZAÇÃO DE DADOS DIRETA
### 1° ETAPA: INSERIR DADOS NA ABA DE HISTÓRICO DE TRANSFERÊNCIAS
<img width="1492" height="413" alt="image" src="https://github.com/user-attachments/assets/11f72d71-f658-4344-9d68-48f495999ec5" />

### 2° ETAPA: APLICAÇÃO DA MACRO E EXECUÇÃO
<img width="1735" height="538" alt="image" src="https://github.com/user-attachments/assets/c7a5ae4a-5a96-42c0-be90-eeb94ee65e97" />


### 3° ETAPA: ATUALIZAÇÃO DA VISUALIZAÇÃO
<img width="1775" height="420" alt="image" src="https://github.com/user-attachments/assets/901286f2-d996-4f87-8ce8-4ea15fa233bf" />

---
Última atualização 16/06/2026
Projeto desenvolvido em período de estágio.


