# Changelog

## v1.5.0 (Versión actual)

### ✨ Nuevas características
- **Centro de Ayuda**: pestaña completa con flujo visual de solicitudes y catálogo detallado de todas las planillas
- **7 nuevos mantenedores**: Precio de Compra (E), Registro Info (D), Cambio Grupo Artículo (R), Cambio Descripción (W), Flete (L), Impuesto Adicional (V), Descuento (K)
- **Total de 11 planillas**: todas las principales de SAP implementadas

### 🔧 Mejoras
- Validación genérica: ahora soporta planillas sin centro (material solamente)
- Mejor detalle en autocompletado: columnas flexibles por planilla
- Cada planilla muestra: columnas, validaciones, ejemplos reales y salida esperada

### 📋 Planillas implementadas
1. Catalogación (C/C1) — habilita material en centro
2. Precio Venta (G) — asigna precio de venta
3. Bloqueo / Desbloqueo (M/N) — bloquea o reactiva compra/venta
4. Libro de Pedido (J) — asigna proveedor y fuente
5. Precio de Compra (E) — carga precio de compra general o por proveedor
6. Registro Info (D) — crea inforecord (asocia material a proveedor)
7. Cambio Grupo Artículo (R) — reasigna grupo de artículo
8. Cambio Descripción (W) — modifica descripción comercial
9. Flete (L) — carga cargo por traslado (ZR01)
10. Impuesto Adicional (V) — asigna impuesto específico (ZR02)
11. Descuento (K) — carga descuento permanente (RA00) o temporal (RA01)

---

## v1.4.0

### ✨ Nuevas características
- **Gestor de Locales**: convertidor de lista vertical a horizontal + mantenedor de clusters personales
- Los usuarios pueden crear clusters (ej: "Prontos RM") y escribir el nombre en la columna Centro
- Clusters se expanden automáticamente al validar

### 🔧 Mejoras
- Separación de Bloqueo y Desbloqueo nuevamente (status vacío = desbloqueo)
- Mejor visualización de filas en la guía de solicitudes (altura uniforme)

---

## v1.3.0

### ✨ Nuevas características
- **Guía de Solicitudes**: 25 tipos de modificación con descripción de planillas requeridas
- Preselección automática de planillas según el tipo elegido
- Distinción entre planillas del portal vs. canal habitual

---

## v1.2.0

### ✨ Nuevas características
- **4 mantenedores principales**: Catalogación, Precio Venta, Bloqueo/Desbloqueo, Libro de Pedido
- Grilla editable con importar/exportar Excel
- Validación de centros, SKU, proveedores contra maestros
- Autocomplete de nombres y descripciones
- Detección de duplicados

### 🔧 Mejoras
- Máscara de fechas automática (DD-MM-AAAA)
- Fecha final por defecto: 31-12-9999
- Expansión de centros por coma: 2003,2549,2626 → una fila por centro

---

## v1.1.0

### ✨ Nuevas características
- **Bandeja de Solicitudes**: historial de envíos con estados (Enviada, En proceso, Aplicada, Rechazada)
- Descarga de Excel procesado por solicitud
- Cambio de estado para equipo Datos Maestros

---

## v1.0.0

### ✨ Características iniciales
- Interfaz Apple-style limpia y moderna
- 4 pestañas: Nueva solicitud, Solicitudes, Gestor de locales, Centro de Ayuda
- Maestros de muestra embebidos (534 centros, 600 SKU, 400 proveedores)
- Carga de Excel maestro completo desde interfaz
- Almacenamiento local (solicitudes + clusters)
