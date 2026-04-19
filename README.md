# ERP Educacional UniFECAF Polo

Aplicação web pronta para Netlify, estruturada como um ERP educacional de polo com base na planilha da sede e overrides locais soberanos.

## Acesso local

Execute:

```powershell
.\start-local-server.cmd
```

Depois abra:

```text
http://localhost:5173
```

Login local de demonstração:

- Usuário: `admin`
- Senha: `admin`

Depois que o Apps Script estiver ativo, o login passa a validar a aba `ERP_Usuarios`.

Usuários iniciais:

| usuario | senha | perfil |
|---|---|---|
| admin | admin | admin |
| financeiro | 123456 | financeiro |
| retencao | 123456 | consultor |
| consultor | 123456 | consultor |

## Estrutura implementada

- Login e layout principal inspirados nas telas de referência: card central, sidebar fixa, cards brancos e Kanban limpo.
- Painel inicial com quatro botões de acesso rápido: Gestão Financeira, Retenção de Alunos, Matrículas e Agendamentos.
- Cabeçalho azul com módulos: Painel, BI, Cadastro, Retenção AVA, Financeiro, Matrículas, Metas, Agendamentos, Avaliações, Fila diária e Segurança.
- Business Intelligence:
  - Metas mensal e anual com barra de progresso.
  - Censo acadêmico por mês/ano, status e gráfico.
  - Saúde financeira por Taxa de Matrícula e Mensalidades Recorrentes, com comparação mensal e anual.
- Janela de Retenção AVA para responsáveis acompanharem acesso e contato.
- Módulo de Inteligência para tomada de decisão de gestão do polo.
- Alertas operacionais de prioridade e falta de contato.
- Persistência operacional em Google Sheets via Apps Script, com fallback local quando a URL ainda não estiver configurada.
- Usuários operacionais na aba `ERP_Usuarios`.
- Acompanhamento de Alunos como Master Data:
  - Nome, CPF mascarado, RA, Curso e Período de Início.
  - Alteração de status diretamente na linha.
  - Override local soberano de status, telefone/WhatsApp, e-mail e status de acompanhamento.
  - Dados locais do polo têm prioridade de exibição sobre a base da sede.
  - Drawer 360° com contato, anotação e baixa de matrícula.
  - Controle de boleto enviado, pagamento de matrícula e isenção.
  - Análise de safra por codificação alfanumérica.
- Retenção AVA:
  - Alerta amarelo para 5 a 7 dias sem acesso.
  - Alerta vermelho para 8 dias ou mais sem acesso.
  - Registro se houve contato com o aluno.
  - Motivo informado, canal, responsável e observação de retenção.
- RBAC:
  - Admin: visão integral.
  - Financeiro: visão financeira operacional.
  - Consultor: valores financeiros bloqueados.
  - Baixa financeira e fluxo de caixa restritos a Admin e Financeiro.
  - Usuários são listados no módulo Segurança e cadastrados na aba `ERP_Usuarios`.
  - Apenas Admin pode alternar a visão de perfil na tela.
- Financeiro:
  - Carteira em dia vs. em atraso.
  - Valor total em atraso.
  - Tabela analítica de inadimplência.
  - Mensalidades da sede em modo leitura.
  - Baixa manual da taxa de matrícula pelo polo.
  - Controle de boleto de matrícula enviado.
  - Controle de alunos isentos de taxa de matrícula.
- Matrículas e CRM de Leads:
  - Catálogo local de cursos.
  - Leads vinculados a curso validado.
  - Kanban de captação com cartões arrastáveis entre etapas.
  - Conversão para Matriculado com criação de registro acadêmico local.
  - Alerta automático de boleto de matrícula pendente.
  - Campo manual de status de pagamento: Pendente, Pago ou Isento.
- Metas:
  - Meta mensal editável.
  - Realizado, gap, conversão e projeção de fechamento.
- Agenda:
  - Docente, assunto, horário, sala e lotação.
  - Bloqueio de conflito de professor e sala.
- Avaliações/TI:
  - Inventário de computadores.
  - Manutenção reduz capacidade disponível.
  - Reserva de provas com bloqueio de overbooking.
- Fila diária:
  - FIFO por horário.
  - Destaque para próximos 30 minutos.
  - Presente, ausente, cancelada/remanejada com justificativa e finalizar.
- Inteligência de gestão:
  - Retenção, risco alto, gap da meta e recursos livres.
  - Sinais automáticos para decisão acadêmica, financeira, comercial e logística.
  - Registro de decisões e plano de ação.

## Usar a planilha como banco operacional

O app já está preparado para criar e usar abas extras na sua planilha:

- `ERP_Config`
- `ERP_Usuarios`
- `ERP_Overrides`
- `ERP_Retencao`
- `ERP_Cursos`
- `ERP_Leads`
- `ERP_Alunos_Local`
- `ERP_Agenda`
- `ERP_Provas`
- `ERP_Arquivo`
- `ERP_Decisoes`
- `ERP_Auditoria`

Como ativar:

1. Abra a planilha Google.
2. Vá em **Extensões > Apps Script**.
3. Cole o conteúdo de [google-apps-script/Code.gs](C:/Users/cleyt/Documents/sofware%20UniFECAF/google-apps-script/Code.gs).
4. Salve o projeto.
5. Execute a função `setupErpTabs` uma vez para criar as abas.
6. Clique em **Implantar > Nova implantação**.
7. Escolha **Aplicativo da Web**.
8. Configure:
   - Executar como: você
   - Quem pode acessar: somente você ou pessoas autorizadas
9. Copie a URL do aplicativo da Web.
10. No Netlify, configure a variável:
    - `ERP_STATE_WEBAPP_URL`

Depois disso, alterações de status, leads, cursos, agenda, provas, metas e decisões passam a ser gravadas na planilha.

## Publicar no Netlify

1. Suba esta pasta para um repositório Git ou use Netlify Drop.
2. O `netlify.toml` já define:
   - pasta pública: `site`
   - funções: `netlify/functions`
3. Para trocar a planilha, configure variáveis no Netlify:
   - `GOOGLE_SHEET_ID`
   - `GOOGLE_SHEET_GID` somente se precisar apontar uma aba específica.
   - `ERP_STATE_WEBAPP_URL` para gravar o estado operacional nas abas ERP.

## Privacidade

A planilha contém dados pessoais. O app mascara CPF na interface, mas os dados ainda são carregados no navegador. Publique em um ambiente com controle de acesso adequado no Netlify e mantenha a planilha com compartilhamento restrito quando possível.

## Validação local

Foram validados:

- Sintaxe dos arquivos JavaScript.
- Carregamento da planilha real via função local.
- Login e renderização do módulo Acompanhamento.
- RBAC do Financeiro.
- Override local de status.
- Painel inicial com botões de acesso rápido.
- Módulo BI com metas mensal/anual, censo acadêmico e saúde financeira.
- Override local de telefone/WhatsApp, e-mail e status de acompanhamento.
- Kanban de matrículas com arrastar e soltar.
- Alerta automático de boleto pendente ao matricular lead.
- Atualização manual de boleto e pagamento de matrícula no CRM.
- Renderização dos 10 módulos.
- Endpoint de estado operacional com fallback local.
