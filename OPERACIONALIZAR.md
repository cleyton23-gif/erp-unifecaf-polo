# Operacionalização do ERP Educacional UniFECAF

Este arquivo é o passo a passo para colocar o aplicativo em operação local, integrar com a planilha Google e publicar no Netlify.

## 1. Arquivos principais

Use estes arquivos como referência:

- `site/index.html`: tela e estrutura do app.
- `site/app.js`: regras do ERP, módulos, filtros, RBAC, retenção, matrícula, agenda e decisões.
- `site/styles.css`: visual do sistema.
- `netlify/functions/sheet.js`: leitura da planilha da sede.
- `netlify/functions/state.js`: gravação/leitura das abas operacionais do ERP.
- `google-apps-script/Code.gs`: script para criar e gravar as abas ERP na planilha Google.
- `netlify.toml`: configuração de deploy do Netlify.
- `start-local-server.cmd`: servidor local para teste.

## 2. Rodar localmente

Abra o PowerShell e execute:

```powershell
cd "C:\Users\cleyt\Documents\sofware UniFECAF"
.\start-local-server.cmd
```

Depois acesse:

```text
http://localhost:5173
```

Login de demonstração:

```text
Usuário: admin
Senha: admin
```

## 3. Testar se o app local está respondendo

Em outro PowerShell:

```powershell
cd "C:\Users\cleyt\Documents\sofware UniFECAF"
Invoke-WebRequest -Uri "http://localhost:5173/" -UseBasicParsing | Select-Object StatusCode
Invoke-WebRequest -Uri "http://localhost:5173/.netlify/functions/sheet" -UseBasicParsing | Select-Object StatusCode
```

Resultado esperado:

```text
StatusCode
----------
200
```

O endpoint de estado operacional pode responder `424` enquanto a URL do Apps Script ainda não estiver configurada. Isso é esperado:

```powershell
try {
  Invoke-WebRequest -Uri "http://localhost:5173/.netlify/functions/state" -UseBasicParsing
} catch {
  $_.Exception.Response.StatusCode
}
```

## 4. Preparar a planilha Google como banco operacional

A planilha da sede continua sendo a base principal da aba Acompanhamento. As alterações do polo ficam em abas ERP separadas.

### 4.1. Abrir Apps Script

1. Abra a planilha Google.
2. Clique em **Extensões**.
3. Clique em **Apps Script**.
4. Apague o conteúdo padrão do editor.
5. Cole todo o conteúdo de:

```text
C:\Users\cleyt\Documents\sofware UniFECAF\google-apps-script\Code.gs
```

6. Salve o projeto.

### 4.2. Criar as abas ERP

No Apps Script:

1. Selecione a função `setupErpTabs`.
2. Clique em **Executar**.
3. Autorize o acesso quando o Google pedir.

O script criará estas abas:

```text
ERP_Config
ERP_Overrides
ERP_Retencao
ERP_Cursos
ERP_Leads
ERP_Alunos_Local
ERP_Agenda
ERP_Provas
ERP_Arquivo
ERP_Decisoes
ERP_Auditoria
```

### 4.3. Criar token de segurança

Crie um token forte. Exemplo no PowerShell:

```powershell
[guid]::NewGuid().ToString("N")
```

Copie o valor gerado.

No Apps Script:

1. Vá em **Configurações do projeto**.
2. Em **Propriedades do script**, clique em **Adicionar propriedade**.
3. Nome:

```text
ERP_STATE_TOKEN
```

4. Valor: cole o token gerado.
5. Salve.

Esse mesmo token será colocado no Netlify.

### 4.4. Publicar o Apps Script como Web App

No Apps Script:

1. Clique em **Implantar**.
2. Clique em **Nova implantação**.
3. Em tipo, escolha **Aplicativo da Web**.
4. Configure:
   - **Executar como:** você.
   - **Quem pode acessar:** qualquer pessoa com o link, ou opção equivalente disponível na sua conta.
5. Clique em **Implantar**.
6. Copie a URL terminada em `/exec`.

Essa URL será usada na variável `ERP_STATE_WEBAPP_URL`.

## 5. Configurar Netlify

O `netlify.toml` já está pronto:

```toml
[build]
  publish = "site"
  functions = "netlify/functions"
```

No Netlify, configure as variáveis de ambiente:

```text
GOOGLE_SHEET_ID=1AoZ9KCNIaIzTEW17MyNFv5O7lVmcoaa8_bFKI8dBY8Q
ERP_STATE_WEBAPP_URL=cole_a_url_do_apps_script_exec
ERP_STATE_TOKEN=cole_o_mesmo_token_do_apps_script
```

Se precisar apontar para uma aba específica da planilha da sede:

```text
GOOGLE_SHEET_GID=gid_da_aba
```

Se a planilha principal exporta corretamente sem GID, deixe `GOOGLE_SHEET_GID` sem configurar.

## 6. Publicar no Netlify

### Opção recomendada: Git conectado ao Netlify

1. Suba esta pasta para GitHub, GitLab ou Bitbucket.
2. No Netlify, clique em **Add new site**.
3. Escolha **Import an existing project**.
4. Conecte o repositório.
5. Confirme:
   - Build command: vazio.
   - Publish directory: `site`.
   - Functions directory: `netlify/functions`.
6. Adicione as variáveis de ambiente da seção anterior.
7. Faça o deploy.

### Opção por CLI

Use esta opção se você tiver Node/npm e Netlify CLI instalados:

```powershell
cd "C:\Users\cleyt\Documents\sofware UniFECAF"
npm install -g netlify-cli
netlify login
netlify init
netlify env:set GOOGLE_SHEET_ID "1AoZ9KCNIaIzTEW17MyNFv5O7lVmcoaa8_bFKI8dBY8Q"
netlify env:set ERP_STATE_WEBAPP_URL "cole_a_url_do_apps_script_exec"
netlify env:set ERP_STATE_TOKEN "cole_o_token"
netlify deploy --prod
```

Se o site já existir no Netlify:

```powershell
cd "C:\Users\cleyt\Documents\sofware UniFECAF"
netlify link
netlify deploy --prod
```

## 7. Validar depois do deploy

Troque `SUA_URL` pela URL do Netlify:

```powershell
Invoke-WebRequest -Uri "https://SUA_URL.netlify.app/" -UseBasicParsing | Select-Object StatusCode
Invoke-WebRequest -Uri "https://SUA_URL.netlify.app/.netlify/functions/sheet" -UseBasicParsing | Select-Object StatusCode
Invoke-WebRequest -Uri "https://SUA_URL.netlify.app/.netlify/functions/state" -UseBasicParsing | Select-Object StatusCode
```

Esperado:

```text
200
```

Se `state` retornar `401`, o token do Netlify está diferente do token do Apps Script.

Se `state` retornar `424`, a variável `ERP_STATE_WEBAPP_URL` não foi configurada.

## 8. Fluxo operacional diário

### 8.1. Atualizar base da sede

1. Extraia os dados da sede.
2. Cole/sobrescreva na aba de origem da planilha.
3. No app, clique em **Sincronizar**.
4. Os dados da sede atualizam, mas os registros locais continuam preservados nas abas ERP.

### 8.2. Retenção AVA

1. Acesse **Retenção AVA**.
2. Verifique:
   - amarelo: 5 a 7 dias sem acesso;
   - vermelho: 8 dias ou mais sem acesso.
3. Clique em **Registrar contato**.
4. Informe se houve contato, canal, motivo, responsável e observação.
5. Salve.

### 8.3. Matrícula do polo

1. Acesse **Acompanhamento** ou **Financeiro**.
2. Abra o aluno/candidato.
3. Marque:
   - boleto enviado;
   - matrícula paga;
   - isento de matrícula, se aplicável.
4. Salve.

### 8.4. Decisão gerencial

1. Acesse **Inteligência**.
2. Leia os sinais automáticos.
3. Registre uma decisão no plano de ação.
4. A decisão será salva em `ERP_Decisoes`.

## 9. Comandos de validação técnica

```powershell
cd "C:\Users\cleyt\Documents\sofware UniFECAF"

& "C:\Users\cleyt\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" --check "site\app.js"
& "C:\Users\cleyt\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" --check "netlify\functions\sheet.js"
& "C:\Users\cleyt\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" --check "netlify\functions\state.js"
& "C:\Users\cleyt\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" --check "tools\local-server.mjs"
```

Validação automatizada no navegador:

```powershell
$env:NODE_PATH="C:\Users\cleyt\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules"
& "C:\Users\cleyt\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" "tools\check-app.mjs"
& "C:\Users\cleyt\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" "tools\check-flows.mjs"
```

## 10. Observações de segurança

- O app mascara CPF, mas os dados continuam sensíveis.
- Use acesso restrito no Netlify.
- Não compartilhe o `ERP_STATE_TOKEN`.
- Se trocar o token, atualize no Apps Script e no Netlify.
- Para produção com muitos usuários, o próximo passo ideal é autenticação real por usuário.

## 11. Fontes oficiais consultadas

- Netlify Docs: https://docs.netlify.com/deploy/create-deploys/
- Netlify CLI: https://docs.netlify.com/api-and-cli-guides/cli-guides/get-started-with-cli
- Google Apps Script Web Apps: https://developers.google.com/apps-script/guides/web
- Apps Script deployments: https://developers.google.com/apps-script/concepts/deployments
