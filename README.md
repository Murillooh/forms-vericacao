# TI · Cadastro de Periféricos (Locagora)

Este projeto contém o código-fonte de um formulário responsivo de **Cadastro de Periféricos** projetado para rodar no **Google Apps Script (GAS)**, persistindo dados diretamente em uma **Planilha do Google (Google Sheets)**.

Possui também um ambiente de desenvolvimento local simulado (mock) para testes rápidos de interface e fluxo de usuário sem necessidade de deploy imediato.

---

## 🛠️ Como Testar Localmente

O projeto utiliza o **Vite** para servir a página de forma leve e rápida na sua máquina de desenvolvimento.

### Requisitos
- **Node.js** instalado (versão 18+ recomendada)

### Passos
1. Abra um terminal na pasta do projeto.
2. Instale as dependências:
   ```bash
   npm install
   ```
3. Inicie o servidor local:
   ```bash
   npm run dev
   ```
4. Acesse a URL fornecida no terminal (normalmente `http://localhost:5173`).
5. **Simulação**: O formulário detecta que está rodando fora do Google Apps Script e carrega dados fictícios no Step 2, permitindo testar a validação do IMEI (15 dígitos), transições de etapas e a mensagem de sucesso ao enviar.

---

## 🚀 Como Implantar no Google Apps Script

Siga estes passos para colocar o formulário em produção conectado à sua planilha corporativa:

### 1. Criar e Configurar a Planilha do Google
1. Crie uma nova Planilha no seu Google Drive (ex: `Cadastro de Equipamentos Corporativos`).
2. Acesse a planilha e adicione/renomeie as seguintes abas:
   - **`Respostas`**: Onde os cadastros serão inseridos. *Nota: Se você não criar, o script criará e formatará esta aba automaticamente na primeira submissão.*
   - **`Opcoes`** (Opcional): Crie duas colunas para listar os modelos disponíveis de notebooks e celulares. 
     - Coluna A: Modelos de Notebooks (cabeçalho na linha 1)
     - Coluna B: Modelos de Celulares (cabeçalho na linha 1)
     - *Se a aba "Opcoes" não existir ou estiver vazia, o sistema usará automaticamente uma lista padrão pré-definida em `Code.gs`.*

### 2. Adicionar o Script à Planilha
1. Com a planilha aberta, clique em **Extensões** > **Apps Script**.
2. No painel à esquerda, você verá um arquivo chamado `Código.gs`. Copie todo o conteúdo do arquivo [Code.gs](file:///c:/Users/Murillo%20Silva/Documents/Projeto%20Forms%20Locagora%20Perifericos/Code.gs) e cole-o substituindo o conteúdo original de `Código.gs`.
3. Salve o arquivo (`Ctrl + S` ou ícone de disquete).
4. Clique em **+ (Adicionar um arquivo)** > **HTML**. Nomeie o arquivo como `index` (o Apps Script adicionará a extensão `.html` automaticamente, ficando `index.html`).
5. Copie todo o conteúdo do arquivo [index.html](file:///c:/Users/Murillo%20Silva/Documents/Projeto%20Forms%20Locagora%20Perifericos/index.html) e cole-o neste novo arquivo criado.
6. Salve o arquivo.

### 3. Publicar como Aplicativo Web
1. No canto superior direito da tela do Apps Script, clique em **Implantar** > **Nova implantação**.
2. Clique na engrenagem de configurações ao lado de "Selecione o tipo" e selecione **Como aplicativo da Web**.
3. Preencha as configurações:
   - **Descrição**: `Cadastro de Periféricos - Versão Inicial`
   - **Executar como**: `Eu` (sua conta de e-mail proprietária da planilha)
   - **Quem tem acesso**: `Qualquer pessoa` (ou `Qualquer pessoa com conta Google` / `Sua Organização` dependendo das políticas da sua empresa)
4. Clique em **Implantar**.
5. Na primeira vez, o Google solicitará autorização de acesso à planilha. Clique em **Autorizar acesso**, faça login com a conta correta, vá em **Avançado** e clique em **Ir para [Nome do Projeto] (não seguro)** e confirme as permissões.
6. Copie a **URL do aplicativo da Web** fornecida. Essa é a URL oficial que você compartilhará com os funcionários para realizar o cadastro.
