# -*- coding: utf-8 -*-
"""Aritmética de números de fila cuando se borran filas de una hoja.

Varias hojas del Master se referencian entre sí **por número de fila**:
`Conciliaciones` guarda `Fila Doc` apuntando a `Facturas`. Borrar una fila
corre todo lo que está debajo, y esas referencias quedan apuntando a otro
documento sin que nada avise.

Vive en su propio módulo para poder probarlo: es una cuenta de una línea, pero
equivocarla desordena vínculos contables en silencio.
"""


def ajustar_referencia(fila: int, borradas) -> int:
    """Dónde queda `fila` después de borrar `borradas`.

    Se le descuenta cuántas filas borradas quedaron POR ENCIMA. El orden de
    `borradas` no importa.

    Preguntar por una fila que se borró es un error de programa, no un caso a
    tolerar: si el llamador cree que sigue existiendo, ya está razonando mal.
    """
    fila = int(fila)
    borradas = {int(b) for b in borradas}
    if fila in borradas:
        raise ValueError("la fila %d se borró: no hay a dónde apuntarla" % fila)
    return fila - sum(1 for b in borradas if b < fila)
