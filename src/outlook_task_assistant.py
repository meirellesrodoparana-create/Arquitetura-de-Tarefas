"""Assistente de tarefas para Outlook (Microsoft Graph).

Fluxo:
1. Autentica via client credentials (Azure AD).
2. Busca e-mails não lidos da Inbox.
3. Classifica prioridade e extrai tarefas.
4. Gera um relatório e envia para Teams/WhatsApp (opcional).

Uso:
    python src/outlook_task_assistant.py
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any
from urllib.parse import urlencode

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


@dataclass
class EmailItem:
    message_id: str
    subject: str
    sender: str
    received_at: str
    body_preview: str


@dataclass
class AnalysisResult:
    priority: str
    reason: str
    confidence: float
    tasks: list[dict[str, Any]]


def _requests_module():
    import requests

    return requests


def require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Variável de ambiente obrigatória ausente: {name}")
    return value


def get_graph_token() -> str:
    tenant_id = require_env("AZURE_TENANT_ID")
    client_id = require_env("AZURE_CLIENT_ID")
    client_secret = require_env("AZURE_CLIENT_SECRET")

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    import msal

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
    )

    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    token = result.get("access_token")
    if not token:
        raise RuntimeError(f"Falha ao obter token: {result}")
    return token


def fetch_unread_emails(token: str, user_email: str, top: int = 10) -> list[EmailItem]:
    requests = _requests_module()
    url = (
        f"{GRAPH_BASE_URL}/users/{user_email}/mailFolders/inbox/messages"
        "?$select=id,subject,from,receivedDateTime,bodyPreview,isRead"
        "&$filter=isRead eq false"
        "&$orderby=receivedDateTime desc"
        f"&$top={top}"
    )

    response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    response.raise_for_status()
    payload = response.json()

    items: list[EmailItem] = []
    for raw in payload.get("value", []):
        items.append(
            EmailItem(
                message_id=raw["id"],
                subject=raw.get("subject", "(sem assunto)"),
                sender=(raw.get("from", {}).get("emailAddress", {}).get("address") or "desconhecido"),
                received_at=raw.get("receivedDateTime", ""),
                body_preview=raw.get("bodyPreview", ""),
            )
        )
    return items


def heuristic_analyze(email: EmailItem) -> AnalysisResult:
    text = f"{email.subject} {email.body_preview}".lower()

    high_markers = ["urgente", "incidente", "produção", "prazo hoje", "bloqueio"]
    medium_markers = ["revisar", "alinhamento", "validação", "pendência"]

    if any(marker in text for marker in high_markers):
        priority = "high"
        reason = "Detectadas palavras-chave críticas no conteúdo."
        confidence = 0.86
    elif any(marker in text for marker in medium_markers):
        priority = "medium"
        reason = "Detectadas palavras-chave de acompanhamento."
        confidence = 0.72
    else:
        priority = "low"
        reason = "Sem marcadores de urgência evidentes."
        confidence = 0.61

    task = {
        "title": f"Tratar e-mail: {email.subject[:90]}",
        "owner": "definir",
        "due_date": datetime.now(timezone.utc).date().isoformat(),
        "source_email_id": email.message_id,
    }

    return AnalysisResult(
        priority=priority,
        reason=reason,
        confidence=confidence,
        tasks=[task],
    )


def send_teams_report(webhook_url: str, markdown_report: str) -> None:
    requests = _requests_module()
    response = requests.post(webhook_url, json={"text": markdown_report}, timeout=30)
    response.raise_for_status()


def send_whatsapp_report(message: str) -> None:
    requests = _requests_module()
    account_sid = require_env("TWILIO_ACCOUNT_SID")
    auth_token = require_env("TWILIO_AUTH_TOKEN")
    from_number = require_env("TWILIO_WHATSAPP_FROM")
    to_number = require_env("WHATSAPP_TO")

    payload = {
        "From": f"whatsapp:{from_number}",
        "To": f"whatsapp:{to_number}",
        "Body": message,
    }
    url = f"https://api.twilio.com/2010-04-01/Accounts/{account_sid}/Messages.json"

    response = requests.post(
        url,
        data=urlencode(payload),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        auth=(account_sid, auth_token),
        timeout=30,
    )
    response.raise_for_status()


def build_report(results: list[tuple[EmailItem, AnalysisResult]]) -> str:
    lines = ["## Relatório de e-mails importantes (Outlook)"]
    high_count = sum(1 for _, res in results if res.priority == "high")
    lines.append(f"- Total analisado: {len(results)}")
    lines.append(f"- Prioridade alta: {high_count}")
    lines.append("")

    for email, analysis in results:
        lines.append(f"### {analysis.priority.upper()} | {email.subject}")
        lines.append(f"- De: {email.sender}")
        lines.append(f"- Recebido em: {email.received_at}")
        lines.append(f"- Motivo: {analysis.reason} (confiança {analysis.confidence:.2f})")
        lines.append(f"- Tarefa sugerida: {analysis.tasks[0]['title']}")
        lines.append("")

    return "\n".join(lines)


def main() -> None:
    user_email = require_env("OUTLOOK_USER_EMAIL")
    token = get_graph_token()

    emails = fetch_unread_emails(token=token, user_email=user_email, top=int(os.getenv("FETCH_TOP", "10")))
    analyses = [(email, heuristic_analyze(email)) for email in emails]

    report = build_report(analyses)
    print(report)

    output_path = os.getenv("REPORT_OUTPUT", "outlook_report.json")
    serializable = [
        {
            "email": email.__dict__,
            "analysis": {
                "priority": analysis.priority,
                "reason": analysis.reason,
                "confidence": analysis.confidence,
                "tasks": analysis.tasks,
            },
        }
        for email, analysis in analyses
    ]
    with open(output_path, "w", encoding="utf-8") as file:
        json.dump(serializable, file, ensure_ascii=False, indent=2)

    teams_webhook = os.getenv("TEAMS_WEBHOOK_URL")
    if teams_webhook:
        send_teams_report(teams_webhook, report)
        print("Relatório enviado ao Teams com sucesso.")

    if os.getenv("ENABLE_WHATSAPP_REPORT", "false").lower() == "true":
        send_whatsapp_report(report)
        print("Relatório enviado no WhatsApp com sucesso.")


if __name__ == "__main__":
    main()
