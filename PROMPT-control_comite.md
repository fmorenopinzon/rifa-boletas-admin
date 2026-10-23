# Oficialización de entrega de Boletas pagadas

Puedes generar una pagina HTML para gestionar la información de entrega de boletas de los asociados al comité del Fondo para oficializar las boletas a jugar. Esta información estará en el Drive como Google Shets con la misma estructura al siguiente Excel RIFA-Sorteo_2026-12-09-control_comite.xlsx. Estos son otros requerimientos:

- El archivo HTML debe llamarse comite.html.
- utilizar otros colores de la hoja de estilos para esta página. Puede ser en tonalidades de un Fondo verde degradado
- La hoja 'Boletas' será la fuente de información de consulta y actualización
- Al abrir el HTML y cargar la información, deberá cargar la lista de los Asociados
- Se deben de tener 2 opciones de consultar y gestionar los pagos de las boletas.
    1. Consultar por número de boleta y cargar los datos que tiene esa boleta, como lo son El asociado asignado, los numeros de la boleta, si ya se ecnuentra registrada como 'Vendida/Pagada' (con el Texto 'Si' o 'No', según el caso) y la 'Fecha Entrega'.
    En caso de que esa boleta NO haya sido marcada como 'Vendida/Pagada', deberá permitir a través de un componente check que se actualice este campo, que en caso de activarse y guardarse deberá marcar la columna en el Excel con una x o X. En el caso de la 'Fecha Entrega' actualizarla automáticamente con la fecha del momento en que se guarde
    2. Consultar por las boletas asignadas al Asociado seleccionado. En este caso, deberá cargar una tabla con los registros de los numeros de la Boleta, si se encuentra 'Vendida/Pagada', 'Fecha Entrega' y en caso de que NO este marcada con una X en la celda 'Vendida/Pagada' cargar un componente chech box para actualizar esta información. Este modo deberá tener un botón de guardado que actualizará la lista de boletas que se hayan sido marcadas como Pagadas y la celda 'Fecha Entrega' debe quedar con la Fecha actual de l operación.
- En otra sección, mostrar la cantidad total de las boletas registradas en la misma fecha que la actual y el valor Total que da al multiplicar esa cantidad de boletas por el Valor de la Boleta (Celda $O$1)
- Igualmente mostrar el valor TOTAL RECAUDADO que se va calculando en el Excel y que corresponde a la celda $G$255

- Registrar el total X Dia como en la Tabla que comienza en la celda I260 donde en la columna I se coloca la fecha de Registro y en la columna J se registra la Formula correspondiente a la suma de valores que coincidan con la Fecha de la columna de la Izquierda