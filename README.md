# Control de Gastos

Aplicación de escritorio para registrar y analizar gastos personales y de negocio. Desarrollada en Python con interfaz gráfica, persistencia en SQLite y soporte para exportar a Excel o CSV.

---

## Características

- Registro de gastos con categoría, descripción, monto y tipo (Personal / Negocio)
- Vista diaria y mensual con totales desglosados
- Análisis por categoría con gráfico de distribución Personal vs. Negocio
- Edición y eliminación de registros existentes
- Exportación a `.xlsx` (con formato) o `.csv`
- Tema oscuro/claro intercambiable en caliente
- Validación robusta de todos los campos antes de tocar la base de datos
- Soporte de formatos de monto en COP (`1.200.000,50`) e inglés (`1,200.50`)

---

## Requisitos

- Python 3.10 o superior
- Las siguientes librerías (ver instalación):

| Librería | Uso |
|---|---|
| `ttkbootstrap` | Interfaz gráfica con temas modernos |
| `openpyxl` | Exportación a Excel (opcional) |

SQLite viene incluido en la librería estándar de Python asi que no requiere instalación adicional.

---

## Instalación

```bash
# Para clonar o descargar el proyecto
git clone https://github.com/tu-usuario/control-gastos.git
cd control-gastos

# Instalar dependencias
pip install ttkbootstrap openpyxl
```

Si no vas a usar la exportación a Excel, `openpyxl` es opcional — la app lo detecta automáticamente.

---

## Uso

```bash
python main.py
```

La base de datos `gastos_andres.db` se crea automáticamente en la misma carpeta del proyecto al primer arranque, tiene asi su nombre debido a que es una App simulando cliente, el cliente ficticio se llama Andrés.

### Registrar un gasto

1. En la pestaña **Hoy**, completa el formulario (monto, categoría, descripción y tipo).
2. Haz clic en **Registrar Gasto**.

El monto acepta varios formatos: `25000`, `25.000`, `25,000`, `1.200.000,50`.

### Editar o eliminar

Selecciona cualquier fila en la tabla y usa los botones **Editar** o **Eliminar** en la parte inferior.

### Vista mensual

Selecciona mes y año, luego haz clic en **Buscar**. Desde ahí puedes exportar el listado a Excel o CSV.

### Categorías

Muestra el total gastado por categoría en el mes seleccionado, con un gráfico de barra que divide Personal y Negocio.

---

## Estructura del proyecto

```
control-gastos/
├── main.py          # Interfaz gráfica (Tkinter + ttkbootstrap)
├── backend.py       # Lógica de negocio y acceso a base de datos
├── test_backend.py  # Tests unitarios (pytest)
└── gastos_andres.db # Base de datos SQLite (se genera automáticamente)
```

**`backend.py`** está dividido en dos clases con responsabilidad única:

- `ValidadorGastos` — valida todos los campos sin tocar la base de datos.
- `GastosDB` — gestiona la persistencia en SQLite, delegando toda validación al validador.

---

## Tests

```bash
pip install pytest
pytest test_backend.py -v
```

Los tests cubren la función de parseo de montos, y todos los métodos de validación: monto, descripción, tipo, fecha e ID, la mayoria de casos posibles para un error estan cubiertos.

---

## Notas

- Compatible con PyInstaller para generar un ejecutable independiente.
- El archivo `.db` se guarda junto al ejecutable, tanto en modo desarrollo como empaquetado.
- Si `openpyxl` no está instalado, la opción de exportar a Excel muestra un aviso y la exportación a CSV sigue funcionando con normalidad, asi que, si no usaras la función de exportación a Excel, simplemente ignoralo.

- Gracias por leer.
