from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx import Document

from rule_engine.embedded_artifacts import run_embedded_artifact_cleanup


class EmbeddedArtifactCleanupTests(unittest.TestCase):
    def test_removes_repeated_footer_paragraphs_and_keeps_body_text(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "sample.docx"
            doc = Document()
            doc.add_paragraph("El operador planta debe informar al supervisor de turno cualquier desviacion critica.")
            for page in range(1, 4):
                doc.add_paragraph("Versión")
                doc.add_paragraph("1")
                doc.add_paragraph('"Este es un documento controlado"')
                doc.add_paragraph("Fecha de Autorización Junio 2023")
                doc.add_paragraph(f"Página {page} de 3")
                doc.add_paragraph("Próxima Revisión Junio 2026")
            doc.save(path)

            records = run_embedded_artifact_cleanup(
                path,
                document_name="sample.docx",
                config={"action": "remove", "enabled": True},
                dry_run=False,
            )

            cleaned = Document(str(path))
            full_text = "\n".join(paragraph.text for paragraph in cleaned.paragraphs)
            self.assertIn("operador planta", full_text)
            self.assertNotIn("documento controlado", full_text)
            self.assertNotIn("Página 2 de 3", full_text)
            self.assertGreaterEqual(sum(1 for record in records if record.applied), 15)

    def test_second_run_is_idempotent(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "sample.docx"
            doc = Document()
            for page in range(1, 3):
                doc.add_paragraph("Versión")
                doc.add_paragraph("1")
                doc.add_paragraph('"Este es un documento controlado"')
                doc.add_paragraph(f"Página {page} de 2")
                doc.add_paragraph("Próxima Revisión Junio 2026")
            doc.save(path)

            run_embedded_artifact_cleanup(path, config={"action": "remove", "enabled": True}, dry_run=False)
            second = run_embedded_artifact_cleanup(path, config={"action": "remove", "enabled": True}, dry_run=False)

            self.assertEqual(0, sum(1 for record in second if record.action in {"remove", "clear_text"}))

    def test_operational_text_is_not_flagged(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "sample.docx"
            doc = Document()
            doc.add_paragraph("Página de procedimiento: el operador debe revisar el estado del equipo antes de iniciar.")
            doc.add_paragraph("Supervisor de Producción debe autorizar la continuidad operacional.")
            doc.save(path)

            records = run_embedded_artifact_cleanup(path, config={"action": "preview", "enabled": True}, dry_run=True)

            self.assertEqual([], records)

    def test_repeated_metadata_tables_are_reviewed_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "sample.docx"
            doc = Document()
            for _ in range(2):
                table = doc.add_table(rows=3, cols=3)
                table.cell(0, 0).text = "Código P-PRPL-OC-506"
                table.cell(0, 1).text = "VERSIÓN"
                table.cell(0, 2).text = "Página 1 de 2"
                table.cell(1, 0).text = "Elaboró:"
                table.cell(1, 1).text = "Revisó:"
                table.cell(1, 2).text = "Aprobó:"
                table.cell(2, 0).text = "Fecha de Autorización Junio 2023"
            doc.save(path)

            records = run_embedded_artifact_cleanup(path, config={"action": "remove", "enabled": True}, dry_run=False)
            table_records = [record for record in records if record.block_type == "table"]

            self.assertEqual(2, len(table_records))
            self.assertIn("protected", {record.action for record in table_records})
            self.assertIn("review", {record.action for record in table_records})


if __name__ == "__main__":
    unittest.main()