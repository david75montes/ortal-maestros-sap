# Portal de Solicitudes de Datos Maestros · SAP

Portal web para que los usuarios soliciten cambios en datos maestros de SAP.
Los solicitantes ingresan datos en una grilla (o importan Excel), el portal los
**valida contra las bases de centros, SKU y proveedores**, autocompleta los
nombres/descripciones, detecta duplicados y genera el archivo procesado listo
para la carga en SAP.

## Características principales

**11 mantenedores implementados:**
- Catalogación (C/C1)
- Precio Venta (G)
- Bloqueo / Desbloqueo (M/N)
- Libro de Pedido (J)
- **Precio de Compra (E)** ← nuevo
- **Registro Info (D)** ← nuevo
- **Cambio Grupo Artículo (R)** ← nuevo
- **Cambio Descripción (W)** ← nuevo
- **Flete (L)** ← nuevo
- **Impuesto Adicional (V)** ← nuevo
- **Descuento (K)** ← nuevo

**Herramientas:**
- **Guía de solicitudes**: 25 tipos de modificación; indica qué planillas enviar
  y preselecciona automáticamente las que existen en el portal.
- **Grilla editable** por planilla, con duplicar fila, importar/exportar Excel,
  máscara de fechas (DD-MM-AAAA) y fecha final 31-12-9999 por defecto.
- **Validación bajo demanda** contra los maestros + detección de filas duplicadas.
- **Expansión de centros**: la columna Centro acepta varios códigos separados por
  coma o el nombre de un *cluster*; se genera una fila por centro.
- **Gestor de locales**: convertidor de lista vertical → horizontal y mantenedor
  de clusters personales (múltiples usuarios, datos privados).
- **Centro de Ayuda**: flujo visual de procesamiento y catálogo detallado de cada
  planilla con columnas, validaciones y ejemplos de uso.
- **Bandeja de solicitudes**: registro de lo enviado, con estados
  (Enviada / En proceso / Aplicada / Rechazada) y descarga de archivos.

## Cómo ejecutarlo en tu computador

Requisito: tener instalado **Node.js 18 o superior** (https://nodejs.org).

1. Abre una terminal dentro de esta carpeta.
2. Instala dependencias:
   ```
   npm install
   ```
3. Levanta el entorno de desarrollo:
   ```
   npm run dev
   ```
4. Abre en el navegador la dirección que muestra la terminal
   (normalmente http://localhost:5173).

Para generar la versión optimizada de producción:
```
npm run build
```
Los archivos quedan en la carpeta `dist/`.

## Estructura

```
portal-maestros-sap/
├── README.md                 ← este archivo
├── NOTAS_TECNICAS.md         ← decisiones, pendientes y puntos de integración
├── package.json
├── vite.config.js
├── index.html
├── src/
│   ├── PortalSAP.jsx          ← componente principal (toda la lógica + UI)
│   └── main.jsx               ← punto de entrada
└── datos/
    └── maestros_muestra.js    ← muestra embebida de centros / SKU / proveedores
```

## Datos maestros

La versión actual trae una **muestra** embebida (534 centros, 600 SKU, 400
proveedores) para poder probar de inmediato. Para validar contra la base
completa, usa el botón **"Cargar Excel maestro"** dentro del portal y selecciona
el Excel que contenga las hojas `CENTROS`, `SKU` y `PROVEEDORES`.

En producción esto se reemplaza por una consulta a la base de datos real
(ver NOTAS_TECNICAS.md).
