# Arquitetura-de-Tarefas

Sim — agora o MVP já suporta reportar por **Outlook + Teams + WhatsApp**.
Ele lê e-mails não lidos via Microsoft Graph, classifica prioridade, sugere tarefas e gera relatório.

## O que foi implementado

- `src/outlook_task_assistant.py`: integração com Outlook + envio opcional para Teams e WhatsApp.
- `.env.example`: variáveis de ambiente necessárias.
- `requirements.txt`: dependências Python.
- `tests/test_outlook_task_assistant.py`: testes automatizados básicos.

## Guia rápido de integração (15–30 min)

### Passo 1 — Registrar o app no Entra ID
1. Acesse o **Microsoft Entra Admin Center**.
2. Entre em **App registrations** → **New registration**.
3. Nome sugerido: `OutlookTaskAssistant`.
4. Vá em **Certificates & secrets** → crie um **Client Secret**.
5. Anote:
   - `Tenant ID`
   - `Client ID`
   - `Client Secret`

### Passo 2 — Permissões no Microsoft Graph
No app criado:
1. **API permissions** → **Add a permission** → **Microsoft Graph**.
2. Selecione **Application permissions**.
3. Adicione `Mail.Read`.
4. Clique em **Grant admin consent**.

> O script usa `client_credentials`, então a permissão precisa ser do tipo **Application**.

### Passo 3 — Configurar seu ambiente local

```bash
cp .env.example .env
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Edite `.env` com suas credenciais Azure e mailbox alvo (`OUTLOOK_USER_EMAIL`).

### Passo 4 — Rodar validação local

```bash
python -m unittest discover -s tests -p 'test_*.py'
python src/outlook_task_assistant.py
```

Se tudo estiver certo, o script:
1. Lê e-mails não lidos (`FETCH_TOP`).
2. Classifica em `high|medium|low`.
3. Sugere tarefa por e-mail.
4. Salva `outlook_report.json`.
5. Publica no Teams/WhatsApp quando habilitado.

---

## Como integrar no seu Outlook (detalhado)

### 1) Criar app no Microsoft Entra ID (Azure)
1. Entre em **Microsoft Entra Admin Center**.
2. Vá em **App registrations** → **New registration**.
3. Nomeie como `OutlookTaskAssistant`.
4. Em **Certificates & secrets**, gere um **Client Secret**.
5. Salve:
   - `Tenant ID`
   - `Client ID`
   - `Client Secret`

### 2) Dar permissões do Microsoft Graph
No app registrado, adicione **Application permissions**:
- `Mail.Read`

Depois clique em **Grant admin consent**.

### 3) Configurar variáveis de ambiente

```bash
cp .env.example .env
```

Preencha os valores no `.env`.

### 4) Instalar dependências

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 5) Executar o assistente

```bash
set -a
source .env
set +a
python src/outlook_task_assistant.py
```

Ao rodar, ele:
1. Busca e-mails **não lidos** da inbox (`FETCH_TOP`).
2. Classifica cada e-mail em `high|medium|low`.
3. Sugere uma tarefa por e-mail.
4. Gera relatório no terminal.
5. Salva saída estruturada em JSON (`REPORT_OUTPUT`).
6. Envia para Teams (se `TEAMS_WEBHOOK_URL` estiver preenchido).
7. Envia para WhatsApp (se `ENABLE_WHATSAPP_REPORT=true`).

## Como reportar via Teams

1. Crie um Incoming Webhook no canal do Teams.
2. Copie a URL e configure:

```env
TEAMS_WEBHOOK_URL=https://...
```

3. Rode o script; ele envia automaticamente quando a variável estiver preenchida.

## Como reportar via WhatsApp

O envio por WhatsApp usa **Twilio WhatsApp API**.

1. Crie conta na Twilio e habilite o canal WhatsApp (sandbox ou número aprovado).
2. Configure no `.env`:

```env
ENABLE_WHATSAPP_REPORT=true
TWILIO_ACCOUNT_SID=ACxxxxxxxx
TWILIO_AUTH_TOKEN=xxxxxxxx
TWILIO_WHATSAPP_FROM=+14155238886
WHATSAPP_TO=+5511999999999
```

3. Rode o script normalmente; o relatório será enviado para o número configurado.

> Dica: no sandbox Twilio, o destinatário precisa entrar no sandbox antes (mensagem de opt-in).

## Estrutura do resultado

Exemplo de saída JSON (arquivo `outlook_report.json`):

```json
[
  {
    "email": {
      "message_id": "AAMk...",
      "subject": "URGENTE: incidente em produção",
      "sender": "cliente@empresa.com",
      "received_at": "2026-02-27T10:14:00Z",
      "body_preview": "Serviço indisponível..."
    },
    "analysis": {
      "priority": "high",
      "reason": "Detectadas palavras-chave críticas no conteúdo.",
      "confidence": 0.86,
      "tasks": [
        {
          "title": "Tratar e-mail: URGENTE: incidente em produção",
          "owner": "definir",
          "due_date": "2026-02-27",
          "source_email_id": "AAMk..."
        }
      ]
    }
  }
]
```

## Troubleshooting rápido

- **401/403 no Graph**: confira `Mail.Read` como Application + admin consent concedido.
- **Mailbox não encontrada**: valide `OUTLOOK_USER_EMAIL` (UPN correto do M365).
- **Sem mensagens no resultado**: a inbox pode não ter e-mails não lidos (`isRead=false`).
- **WhatsApp não envia**: no sandbox Twilio, confirme opt-in do número em `WHATSAPP_TO`.
- **Teams não envia**: revise URL do webhook e permissões do canal.

## Testes

```bash
python -m unittest discover -s tests -p 'test_*.py'
```

## Próximos passos recomendados

Se você quiser, no próximo passo eu também posso evoluir esse MVP para:
- classificação com LLM (em vez de heurística),
- criação automática de tarefas no Planner/Jira,
- execução agendada (cron/GitHub Actions/Azure Functions),
- dashboard web para acompanhamento de SLA.
