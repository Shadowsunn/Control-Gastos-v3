from __future__ import annotations
from datetime import date
from decimal import Decimal, ROUND_HALF_UP  # ← debe estar esta línea
import sqlite3, re, os, sys

# FUNCIONES AUXILIARES (Sueltas sin Formato POO)

def _base_path() -> str:
    """Retorna la carpeta raiz del proyecto, compatible con PyInstaller"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# Función auxiliar de apoyo, no tocar   
def _parsear_monto(monto: int | float | str) -> Decimal:
    """
    Convierte un string de monto (COP o inglés) a float.
    Lanza ValueError si el formato no es reconocible.
    """
    monto_str = str(monto).strip()
    if not re.match(r'^[\d.,]+$', monto_str):
        raise ValueError(f"El monto '{monto}' debe ser un número válido.")

    tiene_punto = '.' in monto_str
    tiene_coma  = ',' in monto_str

    if tiene_punto and tiene_coma:
        if monto_str.rfind(',') > monto_str.rfind('.'):
            limpio = monto_str.replace('.', '').replace(',', '.')
        else:
            limpio = monto_str.replace(',', '')
    elif tiene_coma and not tiene_punto:
        partes = monto_str.split(',')
        if len(partes) > 2:
            limpio = monto_str.replace(',', '')
        else:
            if len(partes[1]) == 3:
                limpio = monto_str.replace(',', '')
            else:
                limpio = monto_str.replace(',', '.')
    elif tiene_punto and not tiene_coma:
        partes = monto_str.split('.')
        if len(partes) > 2:
            limpio = monto_str.replace('.', '')
        elif len(partes[1]) <= 2:
            limpio = monto_str
        else:
            limpio = monto_str.replace('.', '')
    else:
        limpio = monto_str

    return Decimal(limpio).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


# CLASE VALIDADORA:

class ValidadorGastos:                                
    """
    Responsabilidad única, esta clase debe validar los datos de un gasto antes de persistirlos para evitar errores.
    No tiene conexión directa con la base de datos.
    """

    MONTO_MAXIMO = Decimal("999999999")
    DESC_MAX_CHARS     = 200
    TIPOS_VALIDOS      = ("Personal", "Negocio")
    CATEGORIAS_DEFAULT = [
        "Compra de repuestos", "Almuerzo", "Transporte",
        "Arriendo local", "Internet", "Servicios públicos",
        "Nómina", "Varios",
    ]

    def validar_descripcion(self, texto: str, nombre_campo: str = "Descripción") -> tuple[bool, str]:
        """Valida que un texto no esté vacío ni supere el límite de caracteres."""
        if not isinstance(texto, str):                
            return False, f"{nombre_campo} debe ser un texto."
        texto = texto.strip()
        if not texto:
            return False, f"{nombre_campo} no puede estar vacía."
        if len(texto) > self.DESC_MAX_CHARS:         
            return False, f"{nombre_campo} no puede superar {self.DESC_MAX_CHARS} caracteres."
        return True, texto

    def validar_tipo(self, tipo: str) -> tuple[bool, str]:
        """Valida que el tipo sea Personal o Negocio."""
        if not isinstance(tipo, str):
            return False, "El tipo debe ser un texto."
        tipo = tipo.strip().capitalize()      
        if tipo not in self.TIPOS_VALIDOS:
            return False, f"Tipo inválido: '{tipo}'. Debe ser 'Personal' o 'Negocio'."
        return True, tipo

    def validar_monto(self, monto: int | float | str) -> tuple[bool, str | Decimal]:
        """Valida y parsea el monto. Retorna el float limpio si es válido."""
        try:                                          
            monto_limpio = _parsear_monto(monto)
        except ValueError as e:
            return False, str(e)
        if monto_limpio <= 0:
            return False, "El monto debe ser mayor a 0."
        if monto_limpio > self.MONTO_MAXIMO:
            return False, f"El monto supera el máximo permitido de ${self.MONTO_MAXIMO:,}."
        return True, monto_limpio

    def validar_fecha(self, fecha: str) -> tuple[bool, str]:
        """Valida que la fecha sea un string en formato YYYY-MM-DD."""
        if not isinstance(fecha, str):
            return False, "La fecha debe ser un texto en formato YYYY-MM-DD."
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', fecha):
            return False, f"La fecha '{fecha}' no es válida. Usa el formato YYYY-MM-DD."
        try:
            date.fromisoformat(fecha)                
        except ValueError:
            return False, f"La fecha '{fecha}' no es válida. Usa el formato YYYY-MM-DD."
        return True, fecha

    def validar_id(self, id_gasto: int) -> tuple[bool, int|str]:           
        """Valida que un ID sea un entero positivo."""
        if isinstance(id_gasto, bool) or not isinstance(id_gasto, int) or id_gasto <= 0:
            return False, "El ID debe ser un entero positivo."
        return True, id_gasto


# CLASE DE BASE DE DATOS

class GastosDB:
    """
    Responsabilidad única, gestionar la persistencia de gastos en SQLite.
    Delega toda validación a ValidadorGastos.
    """

    def __init__(self) -> None:
        self.db_path   = os.path.join(_base_path(), "gastos_andres.db")
        self.validador = ValidadorGastos()

    def _conectar(self) -> sqlite3.Connection:
        """Retorna una conexión activa a la base de datos."""
        return sqlite3.connect(self.db_path)

    def inicializar_db(self) -> tuple[bool, str]:
        """Crea la tabla gastos si no existe. Seguro para llamar múltiples veces."""
        try:
            with self._conectar() as conn:
                conn.execute("""
                    CREATE TABLE IF NOT EXISTS gastos (
                        id          INTEGER PRIMARY KEY AUTOINCREMENT,
                        fecha       TEXT    NOT NULL,
                        descripcion TEXT    NOT NULL,
                        categoria   TEXT    NOT NULL,
                        monto       REAL    NOT NULL,
                        tipo        TEXT    NOT NULL CHECK(tipo IN ('Personal', 'Negocio'))
                    )
                """)
                conn.commit()
            return True, "Base de datos inicializada correctamente."
        except sqlite3.Error as e:
            return False, f"Error en la base de datos: {e}"

    def agregar_gasto(self, descripcion: str, categoria: str, monto: int|float|str, tipo: str) -> tuple[bool, str]:
        """Inserta un gasto nuevo con la fecha de hoy."""
        ok, descripcion = self.validador.validar_descripcion(descripcion, "Descripción")
        if not ok: return False, descripcion

        ok, categoria = self.validador.validar_descripcion(categoria, "Categoría")
        if not ok: return False, categoria

        ok, monto = self.validador.validar_monto(monto)
        if not ok: return False, monto

        ok, tipo = self.validador.validar_tipo(tipo)
        if not ok: return False, tipo

        try:
            with self._conectar() as conn:
                conn.execute(
                    "INSERT INTO gastos (fecha, descripcion, categoria, monto, tipo) VALUES (?, ?, ?, ?, ?)",
                    (date.today().isoformat(), descripcion, categoria, monto, tipo)
                )
                conn.commit()
            return True, "Gasto registrado correctamente."
        except sqlite3.Error as e:
            return False, f"Error en la base de datos: {e}"

    def obtener_gastos_dia(self, fecha: str|None = None) -> tuple[bool, str|dict]:
        """Retorna los gastos de un día. Si fecha es None, usa hoy."""
        if fecha is None:
            fecha = date.today().isoformat()
        else:
            ok, fecha = self.validador.validar_fecha(fecha)
            if not ok: return False, fecha

        try:
            with self._conectar() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT id, fecha, descripcion, categoria, monto, tipo FROM gastos WHERE fecha = ? ORDER BY id ASC",
                    (fecha,)
                )
                filas = cursor.fetchall()
            return True, {"data": filas, "count": len(filas)}
        except sqlite3.Error as e:
            return False, f"Error en la base de datos: {e}"

    def obtener_gastos_mes(self, mes: int, año: int) -> tuple[bool, str|dict]:
        """Retorna todos los gastos de un mes y año"""
        if isinstance(mes, bool) or isinstance(año, bool) or not isinstance(mes, int) or not isinstance(año, int):
            return False, "El mes y el año deben ser números enteros."
        if not (1 <= mes <= 12):
            return False, f"Mes inválido ({mes}). Debe estar entre 1 y 12."
        if not (2000 <= año <= 2060):
            return False, f"Año inválido ({año}). Debe estar entre 2000 y 2060."

        patron = f"{año}-{str(mes).zfill(2)}-%"
        try:
            with self._conectar() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT id, fecha, descripcion, categoria, monto, tipo FROM gastos WHERE fecha LIKE ? ORDER BY fecha ASC, id ASC",
                    (patron,)
                )
                filas = cursor.fetchall()
            return True, {"data": filas, "count": len(filas)}
        except sqlite3.Error as e:
            return False, f"Error en la base de datos: {e}"

    def obtener_por_categoria(self, mes: int, año: int) -> tuple[bool, str|dict]:
        """
        Retorna el total por categoría y tipo de un mes, ordenado de mayor a menor.
        Cada fila tiene formato: (categoria, tipo, total)
        Una misma categoría puede aparecer dos veces si tiene gastos Personal y Negocio.
        """
        if isinstance(mes, bool) or isinstance(año, bool) or not isinstance(mes, int) or not isinstance(año, int):
            return False, "El mes y el año deben ser números enteros."
        if not (1 <= mes <= 12):
            return False, f"Mes inválido ({mes}). Debe estar entre 1 y 12."
        if not (2000 <= año <= 2060):
            return False, f"Año inválido ({año}). Debe estar entre 2000 y 2060."

        patron = f"{año}-{str(mes).zfill(2)}-%"
        try:
            with self._conectar() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT categoria, tipo, SUM(monto) FROM gastos WHERE fecha LIKE ? GROUP BY categoria, tipo ORDER BY SUM(monto) DESC",
                    (patron,)
                )
                filas = cursor.fetchall()
            return True, {"data": filas, "count": len(filas)}
        except sqlite3.Error as e:
            return False, f"Error en la base de datos: {e}"

    def obtener_año_minimo(self) -> int:
        """Retorna el año más antiguo con gastos registrados. Si no hay datos, retorna el año actual."""
        try:
            with self._conectar() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT MIN(strftime('%Y', fecha)) FROM gastos")
                resultado = cursor.fetchone()[0]
            return int(resultado) if resultado else date.today().year
        except sqlite3.Error:
            return date.today().year

    def editar_gasto(self, id_gasto: int, fecha: str, categoria: str, descripcion: str, monto: int|float|str, tipo: str) -> tuple[bool, str]:
        """Actualiza todos los campos de un gasto existente."""
        ok, id_gasto    = self.validador.validar_id(id_gasto)
        if not ok: return False, id_gasto

        ok, fecha       = self.validador.validar_fecha(fecha)
        if not ok: return False, fecha

        ok, descripcion = self.validador.validar_descripcion(descripcion, "Descripción")
        if not ok: return False, descripcion

        ok, categoria   = self.validador.validar_descripcion(categoria, "Categoría")
        if not ok: return False, categoria

        ok, monto       = self.validador.validar_monto(monto)
        if not ok: return False, monto

        ok, tipo        = self.validador.validar_tipo(tipo)
        if not ok: return False, tipo

        try:
            with self._conectar() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE gastos SET fecha=?, categoria=?, descripcion=?, monto=?, tipo=? WHERE id=?",
                    (fecha, categoria, descripcion, monto, tipo, id_gasto)
                )
                conn.commit()
                if cursor.rowcount == 0:
                    return False, f"No se encontró ningún gasto con ID {id_gasto}."
            return True, "Gasto actualizado correctamente."
        except sqlite3.Error as e:
            return False, f"Error en la base de datos: {e}"

    def eliminar_gasto(self, id_gasto: int) -> tuple[bool, str]:
        """Elimina un gasto por su ID."""
        ok, id_gasto = self.validador.validar_id(id_gasto)
        if not ok: return False, id_gasto

        try:
            with self._conectar() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM gastos WHERE id = ?", (id_gasto,))
                conn.commit()
                if cursor.rowcount == 0:
                    return False, f"No se encontró ningún gasto con ID {id_gasto}."
            return True, "Gasto eliminado correctamente."
        except sqlite3.Error as e:
            return False, f"Error en la base de datos: {e}"