# 🚀 DJConversor — Sistema de Importação e Conversão de Dados (Implantação DJSystem)

**DJConversor** é uma ferramenta desenvolvida especialmente para o **setor de Implantação da DJSystem**, com o objetivo de **converter e importar dados de sistemas legados** (como **SysLoja - Unitak** ou **System - DJSystem antigo**) para o banco de dados Firebird do **DJMonitor / DJPDV**.

O sistema permite importar de forma automatizada **produtos, grupos, marcas, tributações, clientes e contas a receber**, a partir de planilhas ou arquivos **DBF** exportados de sistemas anteriores, garantindo integridade, logs e praticidade no processo de migração.

---

## 🧩 Principais Funcionalidades

**Importação de Produtos Completa**
- Leitura automática de arquivos `.DBF` ou planilhas convertidas.
- Mapeamento dinâmico de colunas (Descrição, Código de Barras, Preço, Estoque, NCM, etc).
- Suporte a **grades e variações** (P/M/G, cores, tamanhos, etc).
- Importação de **códigos alternativos** de barras.
- Criação automática de **grupos, marcas e tributação ICMS**.

**Importação de Clientes**
- Mapeamento dos campos essenciais: Nome, CPF/CNPJ, Tipo de Pessoa e Cidade.
- Configuração inteligente de **Contribuinte ICMS** (manual, automático ou por coluna específica).
- Tratamento automático de acentuação e caracteres especiais.

**Importação de Contas a Receber**
- Conversão de planilhas de contas de clientes.
- Leitura de campos como nome, valor, vencimento, data de caixa, juros, atraso e status.
- Gravação direta na base Firebird.

**Controle Total de Logs e Erros**
- Geração de arquivos `.txt` de logs com datas e mensagens detalhadas.
- Registros de erros, truncamentos e informações gerais durante a importação.
- Indicadores visuais de progresso e mensagens de status em tempo real.

---

## ⚙️ Estrutura do Projeto

- **uImportadorBase.pas**  
  Classe principal (`TImportadorBase`) responsável por toda a lógica de importação, logs, validações, SQL e manipulação de dados Firebird.

- **uImportarProdutos.pas**  
  Interface de importação de produtos, herdando a estrutura base e permitindo selecionar arquivos DBF, mapear colunas e importar produtos.

- **uImportarClientes.pas**  
  Interface dedicada à importação de clientes, com opções de contribuinte ICMS, campos personalizáveis e validação de dados.

- **uImportarContas.pas**  
  (Opcional) Módulo responsável pela importação de contas a receber.

---

## 🧠 Como Utilizar

1. **Abra o programa DJConversor**  
   Ao iniciar, escolha Configuração para configurar a porta do firebird para a conexão com o banco de dados.
   
   <img width="470" height="426" alt="image" src="https://github.com/user-attachments/assets/2458b52f-5c00-4381-ae4a-32d89fb8b46a" />
   <img width="463" height="389" alt="image" src="https://github.com/user-attachments/assets/0b813937-bfb4-46d5-ad34-6e5382a9d746" />

### Importação de Produtos
<img width="1102" height="829" alt="image" src="https://github.com/user-attachments/assets/259e33a8-afa2-49f3-b19b-b5ea7b1b9192" />



3. **Selecione os arquivos**  
   - Clique em **📂 Buscar** para escolher o arquivo `.DBF` exportado do sistema antigo.  
   - Clique em **📁 Buscar** para selecionar o banco `.FDB` do DJMonitor ou DJPDV de destino.

4. **Mapeie as colunas**  
   - Em cada campo (Descrição, Código de Barras, Preço, Grupo, Marca, etc.), escolha a coluna correspondente do DBF.  
   - O mapeamento é flexível e pode variar conforme o sistema de origem.

5. **Configure as opções**  
   - Defina se deseja importar **grades**, **estoque**, ou **códigos alternativos**.

   ### Importação de Clientes
  <img width="854" height="633" alt="image" src="https://github.com/user-attachments/assets/6639168f-388a-4c33-a8b2-26139f61a1f8" />
 
   - No caso de clientes, escolha como determinar o **Contribuinte ICMS**:
     - Todos como **não contribuintes (9)**  
     - Todos como **contribuintes (1)**  
     - Todos como **isentos (2)**  
     - **Automático** (CNPJ = contribuinte, CPF = não contribuinte)  
     - Ou usar a coluna específica do DBF
   ### Importação de Contas a Receber
  <img width="902" height="678" alt="image" src="https://github.com/user-attachments/assets/d5a242a7-a3d8-46ab-9064-246caa05089f" />

  Importe a planilha de contas a receber e indique as respectivas colunas, o "Código do Plano" padrão é o AP (A Prazo) e o código da carteira padrão é "CA".

7. **Inicie a importação 🚀**  
   - Clique em **Importar** e acompanhe a barra de progresso.  
   - O sistema exibirá mensagens de status e salvará logs detalhados na pasta do executável.

---

## 📊 Logs Gerados

O sistema cria automaticamente arquivos de log para auditoria e depuração:

| Arquivo | Descrição |
|----------|------------|
| `log_info.txt` | Informações gerais do processo |
| `log_erros.txt` | Erros críticos durante a importação |
| `log_avisos.txt` | Avisos e duplicidades ignoradas |
| `log_truncados.txt` | Campos truncados por limite de tamanho |
| `log_erros_detalhados.txt` | Linha e motivo de cada erro detectado |

---

## 🔒 Requisitos Técnicos

- **Linguagem:** Free Pascal / Lazarus  
- **Banco de Dados:** Firebird SQL 3.0+  
- **Charset:** UTF-8  
- **Sistemas de Origem Testados:**  
  - SysLoja (Unitak)  
  - System (DJSystem antigo)  
  - Outros sistemas compatíveis com exportação `.DBF`

---

## 🧱 Estrutura Lógica (Simplificada)

```mermaid
graph TD
  A[Arquivo DBF] --> B[Mapeamento de Colunas]
  B --> C[TImportadorBase]
  C --> D[Validação e Sanitização]
  D --> E[Inserção no Firebird]
  E --> F[Logs e Progresso]
