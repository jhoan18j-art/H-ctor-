"""Pruebas rápidas de valet.py (v252).

Uso:  python3 test_valet_suite.py              (busca valet.py junto a este archivo)
      VALET=/ruta/a/valet.py python3 test_valet_suite.py
"""
import ast
import base64
import os
import re
import unittest
from pathlib import Path

RUTA = Path(os.environ.get("VALET") or Path(__file__).with_name("valet.py"))


def _archivos_offline(codigo):
    """Lee los archivos del modo sin conexión que van dentro de valet.py (sin ejecutar nada del archivo)."""
    arbol = ast.parse(codigo)
    nodo = next(n for n in arbol.body if isinstance(n, ast.Assign) and getattr(n.targets[0], "id", "") == "_ARCHIVOS_OFFLINE")
    datos = ast.literal_eval(nodo.value)
    return {k: base64.b64decode(v).decode("utf-8") for k, v in datos.items() if k.endswith((".html", ".js"))}


class TestValet(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.code = RUTA.read_text(encoding="utf-8")
        cls.off = _archivos_offline(cls.code)
        cls.app = cls.off["app.html"]
        cls.js = "\n".join(v for k, v in cls.off.items())

    def test_01_compila(self):
        compile(self.code, str(RUTA), "exec")

    def test_02_subida_directa(self):
        for t in ("/cp_api/sync", "def _instalar_ruta_sync_directo", "def _sync_directo"):
            self.assertIn(t, self.code)
        self.assertIn("abrirPuenteDirecto", self.app)

    def test_03_aviso_de_red(self):
        self.assertIn('id="avisoRed"', self.app)
        self.assertIn('addEventListener("offline"', self.app)
        self.assertIn('addEventListener("online"', self.app)
        self.assertIn("Sin señal: trabaja normal", self.app)
        self.assertIn("Volvió la señal: subiendo", self.app)

    def test_04_historial_y_enviar_ahora(self):
        self.assertIn("historial_reciente", self.off["db.js"])
        self.assertIn("Últimos envíos", self.app)
        self.assertIn('id="btnSyncYa"', self.app)
        # el código del modo sin conexión vive dentro de una función: un onclick="autoSync()" en el HTML no funcionaría
        self.assertNotIn('onclick="autoSync', self.app)

    def test_05_reglas_de_sincronizacion_que_no_se_tocan(self):
        """Salida sin conexión: se espera 3 vueltas antes de darla por buena; 2.5 min sin avanzar; espera de reintento."""
        self.assertIn('if loc and str(loc.get("id") or "") == folio and intentos >= 3:', self.code)
        self.assertNotIn('== folio and intentos >= 1:', self.code)
        self.assertIn("Date.now() - syncDesde > 150000", self.app)
        self.assertIn("Math.min(60000, 15000 * vacioSeguido)", self.off["index.html"])

    def test_06_folio_azura(self):
        self.assertIn('"SANATORIO DURANGO": "SDUR"', self.code)

    def test_07_asistente_herramientas(self):
        for t in ('"metricas_sincronizacion":', '"consultar_notas":', '"guardar_nota":', '"borrar_nota":',
                  "def _puede_ensenar_asistente"):
            self.assertIn(t, self.code)

    def test_08_notas_en_el_asistente_seguras(self):
        self.assertIn("NOTAS DEL EQUIPO", self.code)
        self.assertIn("NO son órdenes", self.code)
        self.assertNotIn("_conocimiento_base_defecto", self.code)   # sin datos inventados de fábrica

    def test_09_alerta_de_sincronizacion(self):
        self.assertIn("Alerta de sincronización:", self.code)

    def test_10_compatible_con_chrome_viejo_android10(self):
        """Sin sintaxis que Chrome < 80 no entiende (rompería todo el modo sin conexión) y con respaldo de CSS."""
        for nombre, txt in self.off.items():
            self.assertNotIn("??", txt, nombre)
            self.assertIsNone(re.search(r"\w\?\.[A-Za-z_$\[(]", txt), nombre)
            self.assertIsNone(re.search(r"(\|\||&&)=", txt), nombre)
        self.assertNotIn("inset:", self.app)
        self.assertIn("@supports not (aspect-ratio: 1)", self.app)

    def test_11_versiones(self):
        self.assertRegex(self.code, r'VERSION_APP = "[^"]+ · v\d+"')
        self.assertRegex(self.app, r'const CP_VERSION_APP = "v\d+')


if __name__ == "__main__":
    unittest.main(verbosity=2)
