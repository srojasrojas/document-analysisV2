from __future__ import annotations

import unittest
from pathlib import Path

from rule_engine.config import load_config
from rule_engine.rules import load_rules


class RuleFixesTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        config = load_config(Path(__file__).resolve().parents[1] / "config.yaml")
        rules = {rule.id: rule for rule in load_rules(config)}
        cls.operador = rules["operador_to_spence"]
        cls.supervisor = rules["supervisor_to_area_execs"]

    def _assert_expansion(self, rule, text: str, expected: str) -> None:
        decision = rule.apply(text)
        self.assertTrue(decision.changed, msg=f"no cambio: {text} [{decision.reason}]")
        self.assertEqual(decision.modified_text, expected)
        second = rule.apply(decision.modified_text)
        self.assertFalse(second.changed, msg=f"no idempotente: {second.modified_text}")

    def _assert_unchanged(self, rule, text: str) -> None:
        decision = rule.apply(text)
        self.assertFalse(decision.changed, msg=f"cambio indebido: {decision.modified_text}")

    # --- operador: descriptores encadenados (issue "en Terreno Zona Autonoma")

    def test_operador_chained_descriptors_not_split(self) -> None:
        self._assert_expansion(
            self.operador,
            "El Operador en Terreno Zona Autónoma debe revisar el equipo.",
            "El Operador en Terreno Zona Autónoma o personal designado por Minera Spence debe revisar el equipo.",
        )

    def test_operador_legacy_split_after_terreno_is_repaired(self) -> None:
        self._assert_expansion(
            self.operador,
            "El Operador de terreno o personal designado por Minera Spence Zona Autónoma debe revisar el equipo.",
            "El Operador de terreno Zona Autónoma o personal designado por Minera Spence debe revisar el equipo.",
        )

    # --- operador: certificado -> calificado

    def test_operador_equipment_target_uses_calificado(self) -> None:
        self._assert_expansion(
            self.operador,
            "El operador de la retroexcavadora debe verificar el área antes de operar.",
            "El operador de la retroexcavadora o personal calificado designado por Minera Spence "
            "debe verificar el área antes de operar.",
        )

    def test_operador_legacy_certificado_is_upgraded(self) -> None:
        self._assert_expansion(
            self.operador,
            "El operador del cargador frontal o personal certificado designado por Minera Spence buscará zona firme.",
            "El operador del cargador frontal o personal calificado designado por Minera Spence buscará zona firme.",
        )

    # --- operador: verbos en futuro antes saltados

    def test_operador_future_tense_verbs_are_marked(self) -> None:
        self._assert_expansion(
            self.operador,
            "El operador verificará que la máquina se encuentre bloqueada.",
            "El operador o personal designado por Minera Spence verificará que la máquina se encuentre bloqueada.",
        )
        self._assert_expansion(
            self.operador,
            "Operador MDC realizará los bloqueos de las válvulas de agua.",
            "Operador MDC o personal designado por Minera Spence realizará los bloqueos de las válvulas de agua.",
        )
        self._assert_expansion(
            self.operador,
            "El operador MDC es el responsable de mantener el aseo del sector.",
            "El operador MDC o personal designado por Minera Spence es el responsable de mantener el aseo del sector.",
        )

    def test_operador_cas_context_still_skipped(self) -> None:
        self._assert_unchanged(self.operador, "El operador CAS debe coordinar la detención.")

    # --- supervisor: no dividir el cargo de su "apellido"

    def test_supervisor_generic_area_not_split(self) -> None:
        self._assert_expansion(
            self.supervisor,
            "El Supervisor de desarrollo debe autorizar el ingreso.",
            "El Supervisor de desarrollo o Ejecutivos del Área debe autorizar el ingreso.",
        )

    def test_supervisor_legacy_split_is_repaired(self) -> None:
        self._assert_expansion(
            self.supervisor,
            "El Supervisor o Ejecutivos del Área de desarrollo debe autorizar el ingreso.",
            "El Supervisor de desarrollo o Ejecutivos del Área debe autorizar el ingreso.",
        )

    def test_supervisor_legacy_glued_split_is_repaired(self) -> None:
        self._assert_expansion(
            self.supervisor,
            "Supervisor o Ejecutivos del ÁreaEjecución procesos Área Húmeda debe coordinar.",
            "Supervisor Ejecución procesos Área Húmeda o Ejecutivos del Área debe coordinar.",
        )

    # --- supervisor: empresa contratista exenta

    def test_supervisor_contractor_company_is_not_modified(self) -> None:
        self._assert_unchanged(
            self.supervisor, "El supervisor de la empresa contratista debe coordinar con el área."
        )
        self._assert_unchanged(
            self.supervisor, "El Supervisor de empresa colaboradora debe revisar los permisos."
        )
        self._assert_unchanged(self.supervisor, "El supervisor contratista informará los hallazgos.")

    def test_supervisor_regular_cases_still_expand(self) -> None:
        self._assert_expansion(
            self.supervisor,
            "El Supervisor de turno debe autorizar el bloqueo.",
            "El Supervisor de turno o Ejecutivos del Área debe autorizar el bloqueo.",
        )


if __name__ == "__main__":
    unittest.main()
