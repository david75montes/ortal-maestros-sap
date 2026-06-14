# Notas técnicas

## Arquitectura actual

Es una aplicación React de un solo componente (`src/PortalSAP.jsx`) que corre
100% en el navegador. No requiere backend para la demo: lee/escribe Excel con
SheetJS (xlsx) y valida contra maestros cargados en memoria.

## Persistencia

El portal usa el almacenamiento del entorno donde está embebido:

- **Solicitudes enviadas**: almacenamiento *compartido* (todos los usuarios del
  portal las ven). Pensado para que el equipo de Datos Maestros las gestione.
- **Clusters de locales**: almacenamiento *personal* (cada usuario ve los suyos).

> En un despliegue propio sin ese almacenamiento, ambos caen a memoria de sesión.
> Para producción, reemplazar por una base de datos real (ver más abajo).

## Maestros

- `datos/maestros_muestra.js` contiene una muestra (centros completos + 600 SKU +
  400 proveedores).
- Los maestros completos (≈49.000 SKU) NO se embeben por peso; se cargan con el
  botón "Cargar Excel maestro", que lee las hojas CENTROS / SKU / PROVEEDORES.

## Puntos de integración para producción

1. **Maestros desde base de datos**: reemplazar el objeto `MAESTROS` y la función
   `cargarMaestro` por llamadas a una API (`GET /centros`, `/skus`, `/proveedores`
   o un endpoint de validación por lote).
2. **Envío de solicitudes**: la función `guardarSolicitud` debe apuntar a un
   backend (`POST /solicitudes`) que guarde el folio, el solicitante, las filas
   validadas y el estado, y que genere/derive el archivo LSMW.
3. **Autenticación y roles**: hoy no hay login. Falta separar:
   - *Solicitante*: crea solicitudes y ve solo las propias.
   - *Equipo Datos Maestros*: ve todas y cambia estados.
4. **Reglas de unicidad por mantenedor** (PENDIENTE): cada planilla tiene una
   propiedad opcional `claveUnica` (lista de campos). Si se omite, ninguna fila
   puede ser idéntica a otra. Falta confirmar la clave real de cada mantenedor.
   Ejemplos:
   ```js
   claveUnica: ["centro", "sku"]                 // un registro por centro+SKU
   claveUnica: ["centro", "sku", "fechaInicio"]  // permite distinto período
   ```

## Mantenedores implementados vs. guía

La guía contempla 25 tipos de modificación y planillas A…X1. Hoy el portal
implementa 4 planillas: Catalogación (C/C1), Precio Venta (G),
Bloqueo/Desbloqueo (M/N) y Libro de Pedido (J). El resto se marca en la guía
como "Canal habitual". Para sumar una nueva planilla basta agregar un objeto al
arreglo `PLANILLAS` con sus `userCols`, `autoCols`, `salida` y `ejemplo`, y
asignar su `portalId` en el mapa `LETRAS`.

## Validaciones incluidas

- Centro / SKU / Proveedor existen en el maestro.
- Fecha con formato DD-MM-AAAA (máscara automática al escribir).
- Precio numérico > 0.
- Status Bloqueo: Z1–Z4 o vacío (desbloqueo).
- Filas duplicadas (tras expandir centros).
- Expansión de clusters y de centros separados por coma.
