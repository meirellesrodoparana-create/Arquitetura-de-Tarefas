import unittest

from src.outlook_task_assistant import AnalysisResult, EmailItem, build_report, heuristic_analyze


class OutlookAssistantTests(unittest.TestCase):
    def test_heuristic_marks_high_for_urgent_subject(self):
        email = EmailItem(
            message_id="1",
            subject="URGENTE: incidente em produção",
            sender="cliente@empresa.com",
            received_at="2026-01-01T00:00:00Z",
            body_preview="serviço fora do ar",
        )

        result = heuristic_analyze(email)

        self.assertEqual(result.priority, "high")
        self.assertGreaterEqual(result.confidence, 0.8)
        self.assertEqual(result.tasks[0]["source_email_id"], "1")

    def test_build_report_includes_summary_and_subject(self):
        email = EmailItem(
            message_id="2",
            subject="Revisar proposta",
            sender="time@empresa.com",
            received_at="2026-01-01T00:00:00Z",
            body_preview="",
        )
        analysis = AnalysisResult(
            priority="medium",
            reason="Detectadas palavras-chave de acompanhamento.",
            confidence=0.72,
            tasks=[{"title": "Tratar e-mail: Revisar proposta"}],
        )

        report = build_report([(email, analysis)])

        self.assertIn("Total analisado: 1", report)
        self.assertIn("Revisar proposta", report)
        self.assertIn("MEDIUM", report)


if __name__ == "__main__":
    unittest.main()
