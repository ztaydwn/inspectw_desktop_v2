# InspectW Desktop

## Descripción

**InspectW Desktop** es una aplicación de escritorio desarrollada en Python con interfaz gráfica (PyQt6) para procesar y analizar datos de inspecciones fotográficas. Permite cargar archivos ZIP o carpetas que contengan fotos de inspecciones, agruparlas automáticamente según descripciones y grupos predefinidos, aplicar recomendaciones inteligentes basadas en datos históricos utilizando un motor de IA, y generar informes profesionales en formatos PPTX (PowerPoint) o XLSX (Excel).

### Funcionalidades Principales
- **Carga de Datos**: Soporta archivos ZIP o carpetas con estructura específica (archivos `descriptions.txt`, `grupos.txt` y fotos en subcarpetas).
- **Procesamiento Inteligente**: Parsea descripciones para agrupar fotos por categorías (ej. "1 Paredes", "2 Techos"), asigna detalles específicos y aplica recomendaciones basadas en histórico.
- **Recomendaciones IA**: Utiliza un motor de recomendaciones (basado en datos históricos en CSV) para sugerir acciones o observaciones para cada grupo.
- **Generación de Informes**:
  - **PPTX**: Informes en formato A4 con diapositivas que incluyen fotos organizadas en cuadrícula, ubicaciones/detalles y recomendaciones.
  - **XLSX**: Informes tabulares con datos estructurados, incluyendo control de documentos opcional.
- **Interfaz Gráfica**: Fácil de usar con botones para cargar datos, limpiar, y exportar informes. Soporta procesamiento en hilos para evitar bloqueos.
- **Compatibilidad**: Diseñado para inspecciones de construcción o similares, con soporte para archivos de control de documentos (JSON, CSV, TXT).

### Casos de Uso
- Inspecciones de obras civiles o edificaciones.
- Análisis de defectos en fotos agrupadas por secciones.
- Generación de reportes automatizados para clientes o supervisores.

## Requerimientos del Sistema

### Hardware
- **Procesador**: Intel i3 o equivalente (recomendado i5+ para procesamiento rápido).
- **Memoria RAM**: Mínimo 4 GB (recomendado 8 GB para archivos grandes).
- **Almacenamiento**: 500 MB libres para la aplicación + espacio para datos (depende del tamaño de ZIPs/fotos).
- **Pantalla**: Resolución mínima 1024x768 para la interfaz.

### Software
- **Sistema Operativo**: Windows 10/11 (probado), Linux (Ubuntu 18+), macOS (10.14+).
- **Python**: Versión 3.6 o superior (recomendado 3.8-3.11 para compatibilidad con PyQt6). Descarga desde [python.org](https://www.python.org/downloads/).
  - Asegúrate de marcar "Add Python to PATH" durante la instalación.
- **Dependencias de Python** (instálalas automáticamente con `pip install -r requirements.txt`):
  - `PyQt6`: Para la interfaz gráfica.
  - `python-pptx`: Para generar informes PPTX.
  - `openpyxl`: Para manejar archivos Excel (XLSX).
  - `pandas`: Para procesamiento de datos CSV y análisis.
- **Opcional**:
  - Node.js: Solo si lo necesitas para otros proyectos; instálalo por separado para evitar conflictos con Python (ver sección de Troubleshooting).

### Archivos del Proyecto
- `app/`: Código fuente principal.
- `datos/`: Archivos de ejemplo (historico.csv, infoproyect.txt, portada.png, etc.).
- `requirements.txt`: Lista de dependencias.
- `test_imports.py`: Script para verificar dependencias.
- `InspectW.spec`: Archivo para empaquetado con PyInstaller.

## Instalación

1. **Instalar Python**:
   - Descarga e instala Python desde [python.org](https://www.python.org/downloads/).
   - Verifica la instalación: Abre una terminal y ejecuta `python --version` (debe mostrar 3.6+).

2. **Clonar o Copiar el Proyecto**:
   - Copia el directorio `inspectw_desktop` a tu máquina local.

3. **Configurar Entorno Virtual (Recomendado)**:
   - Navega al directorio del proyecto: `cd inspectw_desktop`.
   - Crea un entorno virtual: `python -m venv inspectw_env`.
   - Actívalo: En Windows: `inspectw_env\Scripts\activate`; en Linux/macOS: `source inspectw_env/bin/activate`.

4. **Instalar Dependencias**:
   - Ejecuta: `pip install -r requirements.txt`.
   - Si hay errores, actualiza pip: `python -m pip install --upgrade pip`.
   - Verifica: Ejecuta `python test_imports.py` para confirmar que todas las bibliotecas se instalaron correctamente.

5. **Ejecutar la Aplicación**:
   - Desde el directorio del proyecto: `python app/main.py`.
   - Se abrirá la interfaz gráfica. Si no se abre, revisa errores en la terminal.

### Notas de Instalación
- Si usas un IDE como VS Code, configura el intérprete de Python al entorno virtual creado.
- Para distribución, usa PyInstaller: `pyinstaller InspectW.spec` para generar un ejecutable independiente.

## Uso

### Interfaz Principal
- **Botón "Cargar ZIP(s)"**: Selecciona uno o más archivos ZIP con datos de inspección.
- **Botón "Cargar Carpeta"**: Selecciona una carpeta con archivos `descriptions.txt`, `grupos.txt` y subcarpetas de fotos.
- **Botón "Cargar historico.csv (opcional)"**: Carga un archivo CSV con datos históricos para recomendaciones.
- **Botón "Limpiar"**: Borra datos cargados.
- **Botón "Generar Informe A4 (PPTX)"**: Crea un informe PPTX con fotos y recomendaciones.
- **Botón "Generar Informe (XLSX)"**: Crea un informe XLSX tabular.

### Estructura de Datos Esperada
- **descriptions.txt**: Archivo de texto con formato `[carpeta] archivo.jpg Description: código detalle` (ej. `[Paredes] foto1.jpg Description: 1 Grietas en pared`).
- **grupos.txt**: Lista de grupos numerados (ej. `1 Paredes`, `2 Techos`).
- **Fotos**: Archivos JPG en subcarpetas correspondientes.
- **historico.csv**: CSV con columnas para recomendaciones históricas (opcional, usado para IA).
- **control_documents**: Archivo opcional (JSON/CSV/TXT) para mapear números a situaciones en informes XLSX.

### Ejemplo de Flujo de Trabajo
1. Carga un ZIP con fotos y archivos de texto.
2. (Opcional) Carga `historico.csv` para recomendaciones.
3. Selecciona un grupo en la lista para ver fotos en miniatura.
4. Genera un informe PPTX o XLSX, elige ubicación de guardado.
5. El progreso se muestra en la barra; el informe se guarda automáticamente.

### Atajos y Consejos
- Procesa múltiples ZIPs a la vez para acumular datos.
- Los informes PPTX incluyen hasta 12 fotos por diapositiva (4x3), con texto generado automáticamente para detalles.
- Cancela procesos largos con el botón "Cancelar" en el diálogo de progreso.

## Troubleshooting

### Problemas Comunes y Soluciones

#### 1. Errores al Instalar Node.js y Python
- **Síntoma**: La aplicación funcionaba antes, pero después de instalar Node.js, lanza errores como "ModuleNotFoundError" o fallos de importación.
- **Causa**: El instalador de Node.js puede instalar o modificar una versión de Python, alterando variables de entorno (`PATH`) o sobrescribiendo paquetes.
- **Solución**:
  - Verifica Python: `python --version` (debe ser 3.6+).
  - Reinstala Python desde cero si es necesario.
  - Instala Node.js por separado, desmarcando opciones de Python en el instalador.
  - En Windows, edita `PATH` para priorizar Python del sistema.
  - Crea un entorno virtual y reinstala dependencias: `pip install -r requirements.txt`.

#### 2. Errores de GUI (PyQt6)
- **Síntoma**: La aplicación no se abre o muestra errores de ventana.
- **Causa**: Problemas con Qt o drivers gráficos.
- **Solución**:
  - Instala dependencias adicionales: `pip install pyqt6-tools`.
  - En Windows, instala Visual C++ Redistributables desde Microsoft.
  - Ejecuta en un entorno con GUI (no servidor headless).

#### 3. Errores de Procesamiento
- **Síntoma**: "Faltan 'descriptions.txt' o 'grupos.txt'".
- **Causa**: Estructura de ZIP/carpeta incorrecta.
- **Solución**: Asegúrate de que los archivos estén en la raíz del ZIP/carpeta.

#### 4. Errores de Recomendaciones
- **Síntoma**: "Error cargando recomendaciones".
- **Causa**: Archivo `historico.csv` malformado o faltante.
- **Solución**: Verifica el CSV (columnas correctas) o omite el histórico.

#### 5. Rendimiento Lento
- **Síntoma**: Procesamiento lento con muchos archivos.
- **Causa**: Archivos grandes o hardware limitado.
- **Solución**: Reduce resolución de fotos o usa máquina más potente.

### Depuración General
- Ejecuta `python test_imports.py` para verificar dependencias.
- Revisa logs en la terminal al ejecutar `python app/main.py`.
- Para soporte, proporciona el mensaje exacto de error y versión de Python/SO.

## Contribución y Soporte

- **Repositorio**: (Agrega enlace si aplica).
- **Issues**: Reporta bugs o solicita features.
- **Licencia**: (Especifica si aplica, ej. MIT).

---

*Este documento se actualiza con el código fuente. Para versiones específicas, revisa el historial de commits.*