"""
nlg_utils.py
=================

Este módulo implementa la lógica de agrupamiento y generación de
oraciones (NLG) basada en reglas descrita en el informe. Está
pensado para ser integrado en programas que manipulan objetos con
descripciones de hallazgos (por ejemplo, inspecciones de
infraestructuras) y la ubicación o variable asociada (por
ejemplo, el nombre de la carpeta o posición donde se encontró el
hallazgo).

La función principal expuesta es ``agrupa_y_redacta``. Esta
función toma como entrada una lista de tuplas ``(descripcion,
variable)`` y devuelve una lista de oraciones en castellano que
resumen los hallazgos agrupados. Internamente utiliza el
algoritmo ``SequenceMatcher`` de la biblioteca estándar de
Python para calcular la similitud entre cadenas y agrupar
descripciones similares según un umbral configurable.

Las oraciones de salida se generan a partir de plantillas
dinámicas. Dependiendo de ciertas palabras clave presentes en la
descripción (por ejemplo, ``"fisura"`` o ``"IG"``), se escoge
una plantilla diferente. Si no hay ninguna coincidencia de
palabras clave, se utiliza una plantilla por defecto.

Uso ejemplo:

.. code-block:: python

    from nlg_utils import agrupa_y_redacta

    entradas = [
        ("Fisura longitudinal de 10 cm", "Viga 1"),
        ("fisura long. de 8 cm", "Viga 2"),
        ("IG: falta de tornillos", "Pilar A"),
        ("IG falta de sujetadores", "Pilar B"),
    ]
    oraciones = agrupa_y_redacta(entradas, umbral_similitud=0.8)
    for o in oraciones:
        print(o)

Salida probable::

    Se encontró fisura longitudinal de 10 cm en Viga 1 y Viga 2
    En Pilar A y Pilar B, IG: falta de tornillos

Esta implementación es deliberadamente sencilla y evita el uso de
modelos de aprendizaje automático. Está pensada para ofrecer
transparencia, control y facilidad de mantenimiento.
"""

from difflib import SequenceMatcher
from typing import Dict, List, Tuple


def _normaliza_descripcion(texto: str) -> str:
    """Preprocesa una descripción para facilitar la comparación.

    Actualmente la normalización se limita a convertir la cadena a
    minúsculas y eliminar espacios duplicados. Se podrían añadir
    otros pasos, como eliminar acentos o puntuación, según la
    naturaleza de los datos.

    Args:
        texto: Cadena original de la descripción.

    Returns:
        Una versión normalizada de la descripción.
    """
    return " ".join(texto.lower().split())


def agrupar_descripciones(
    entradas: List[Tuple[str, str]], umbral_similitud: float = 0.8
) -> List[Dict[str, List[str]]]:
    """Agrupa descripciones similares según un umbral de similitud.

    Cada entrada es una tupla ``(descripcion, variable)``. La
    función agrupa las descripciones que alcanzan o superan el
    umbral de similitud utilizando ``SequenceMatcher.ratio``. La
    primera descripción de un grupo se toma como referencia.

    Args:
        entradas: Una lista de pares donde el primer elemento es la
            descripción del hallazgo y el segundo la variable
            asociada (por ejemplo, la ubicación).
        umbral_similitud: Valor entre 0 y 1 que determina a partir
            de qué similitud dos descripciones se consideran del
            mismo grupo. Un valor más alto resulta en grupos más
            estrictos.

    Returns:
        Una lista de diccionarios. Cada diccionario tiene las
        llaves ``"descripcion"`` (la descripción representativa del
        grupo) y ``"variables"`` (lista de variables asociadas a
        las entradas del grupo).
    """
    grupos: List[Dict[str, List[str]]] = []
    for descripcion, variable in entradas:
        desc_norm = _normaliza_descripcion(descripcion)
        asignado = False
        for grupo in grupos:
            referencia_norm = _normaliza_descripcion(grupo["descripcion"])
            similitud = SequenceMatcher(None, desc_norm, referencia_norm).ratio()
            if similitud >= umbral_similitud:
                grupo["variables"].append(variable)
                asignado = True
                break
        if not asignado:
            grupos.append({"descripcion": descripcion, "variables": [variable]})
    return grupos


def _seleccionar_plantilla(descripcion: str) -> str:
    """Devuelve una plantilla de oración según la descripción.

    Anteriormente, esta función seleccionaba diferentes plantillas
    basadas en palabras clave. Ahora, para cumplir con el requisito
    de un formato consistente, siempre devuelve la misma plantilla:

    .. code-block:: text

        "{descripcion} en {variables}"

    Esto asegura que todas las oraciones generadas sigan la
    estructura de "descripción" seguida de "en" y las "variables"
    asociadas.

    Args:
        descripcion: Cadena de descripción representativa del grupo.

    Returns:
        Una cadena de plantilla con los marcadores ``{descripcion}``
        y ``{variables}`` listos para formatear.
    """
    # Se unifica la plantilla para seguir siempre el formato "descripción en variables".
    return "{descripcion} en {variables}"


def _formatear_variables(variables: List[str]) -> str:
    """Devuelve una cadena con las variables separadas adecuadamente.

    Si hay más de una variable, se separan con comas y la última
    precedida de "y" para un lenguaje natural más fluido. Por
    ejemplo, ``["A", "B", "C"]`` se convierte en ``"A, B y C"``.

    Args:
        variables: Lista de variables (ubicaciones) asociadas al
            grupo.

    Returns:
        Cadena con las variables unidas.
    """
    if not variables:
        return ""
    variables_unicas = []
    # Preservamos el orden de aparición pero evitamos duplicados
    seen = set()
    for v in variables:
        if v not in seen:
            variables_unicas.append(v)
            seen.add(v)
    if len(variables_unicas) == 1:
        return variables_unicas[0]
    if len(variables_unicas) == 2:
        return " y ".join(variables_unicas)
    return ", ".join(variables_unicas[:-1]) + " y " + variables_unicas[-1]


def redactar_oracion(grupo: Dict[str, List[str]]) -> str:
    """Construye una oración final para un grupo de descripciones.

    Args:
        grupo: Diccionario con claves ``"descripcion"`` y
            ``"variables"`` generado por ``agrupar_descripciones``.

    Returns:
        Una cadena con la oración generada, con la descripción y
        variables insertadas en la plantilla seleccionada.
    """
    descripcion = grupo["descripcion"]
    variables = grupo["variables"]
    plantilla = _seleccionar_plantilla(descripcion)
    vars_formateadas = _formatear_variables(variables)
    oracion = plantilla.format(descripcion=descripcion, variables=vars_formateadas)

    # Capitalizar solo la primera letra de la oración resultante.
    if oracion:
        return oracion[0].upper() + oracion[1:]
    return ""


def agrupa_y_redacta(
    entradas: List[Tuple[str, str]], umbral_similitud: float = 0.8
) -> List[str]:
    """Agrupa entradas similares y genera oraciones para cada grupo.

    Esta función es el punto de entrada más conveniente cuando se
    dispone de una lista de tuplas ``(descripcion, variable)``. Se
    encarga de:

    1. Agrupar las descripciones que superan el umbral de similitud.
    2. Seleccionar una plantilla adecuada para cada grupo según las
       palabras clave de la descripción representativa.
    3. Formatear las variables de forma natural (con comas y "y").
    4. Rellenar la plantilla con la descripción y las variables.

    Args:
        entradas: Lista de tuplas con la descripción del hallazgo y
            la variable asociada.
        umbral_similitud: Umbral de similitud entre 0 y 1 para
            agrupar descripciones. El valor por defecto es 0.8,
            recomendado en el informe. Ajuste este valor según
            la tolerancia a variaciones de su conjunto de datos.

    Returns:
        Una lista de oraciones en castellano que resumen los
        hallazgos agrupados.
    """
    grupos = agrupar_descripciones(entradas, umbral_similitud)
    return [redactar_oracion(gr) for gr in grupos]
