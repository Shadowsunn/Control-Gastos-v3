from __future__ import annotations
from decimal import Decimal
import pytest
from backend import _parsear_monto, ValidadorGastos

v = ValidadorGastos()

# _parsear_monto
class TestParsearMonto:
    def test_entero_simple(self):
        assert _parsear_monto("1200") == Decimal("1200.00")

    def test_decimal_punto(self):
        assert _parsear_monto("1200.50") == Decimal("1200.50")

    def test_decimal_coma(self):
        assert _parsear_monto("1200,50") == Decimal("1200.50")

    def test_miles_punto_cop(self):
        assert _parsear_monto("1.200.000") == Decimal("1200000.00")

    def test_miles_coma_ingles(self):
        assert _parsear_monto("1,200") == Decimal("1200.00")

    def test_cop_completo(self):
        assert _parsear_monto("1.200.000,50") == Decimal("1200000.50")

    def test_invalido_lanza_error(self):
        with pytest.raises(ValueError):
            _parsear_monto("abc")


# ValidadorGastos.validar_monto
class TestValidarMonto:
    def test_monto_valido(self):
        ok, val = v.validar_monto("25000")
        assert ok is True
        assert val == Decimal("25000.00")

    def test_monto_cero(self):
        ok, msg = v.validar_monto("0")
        assert ok is False
        assert "mayor a 0" in msg

    def test_monto_negativo(self):
        ok, msg = v.validar_monto("-100")
        assert ok is False

    def test_monto_supera_maximo(self):
        ok, msg = v.validar_monto("9999999999")
        assert ok is False
        assert "máximo" in msg


# ValidadorGastos.validar_descripcion
class TestValidarDescripcion:
    def test_valida(self):
        ok, val = v.validar_descripcion("Almuerzo")
        assert ok is True
        assert val == "Almuerzo"

    def test_vacia(self):
        ok, msg = v.validar_descripcion("   ")
        assert ok is False
        assert "vacía" in msg

    def test_supera_limite(self):
        ok, msg = v.validar_descripcion("x" * 201)
        assert ok is False

    def test_no_string(self):
        ok, msg = v.validar_descripcion(123)
        assert ok is False


# ValidadorGastos.validar_tipo
class TestValidarTipo:
    def test_personal(self):
        ok, val = v.validar_tipo("Personal")
        assert ok is True

    def test_negocio_minuscula(self):
        ok, val = v.validar_tipo("negocio")
        assert ok is True
        assert val == "Negocio"

    def test_invalido(self):
        ok, msg = v.validar_tipo("Familiar")
        assert ok is False


# ValidadorGastos.validar_fecha
class TestValidarFecha:
    def test_valida(self):
        ok, val = v.validar_fecha("2026-05-27")
        assert ok is True

    def test_formato_sin_guiones(self):
        ok, msg = v.validar_fecha("20260527")
        assert ok is False

    def test_fecha_imposible(self):
        ok, msg = v.validar_fecha("2026-13-45")
        assert ok is False


# ValidadorGastos.validar_id
class TestValidarId:
    def test_valido(self):
        ok, val = v.validar_id(1)
        assert ok is True
        assert val == 1

    def test_cero(self):
        ok, msg = v.validar_id(0)
        assert ok is False

    def test_negativo(self):
        ok, msg = v.validar_id(-5)
        assert ok is False

    def test_booleano(self):
        ok, msg = v.validar_id(True)
        assert ok is False