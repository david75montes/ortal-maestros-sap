# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Portal de Solicitudes de Datos Maestros · SAP** is a React web application that allows users to request changes to SAP master data. The application runs entirely in the browser (100% client-side), validates against master data (centers, SKUs, suppliers), auto-completes names/descriptions, detects duplicates, and generates processed Excel files ready for SAP upload.

- **11 implemented master data modules**: Catalogación (C/C1), Precio Venta (G), Bloqueo/Desbloqueo (M/N), Libro de Pedido (J), Precio de Compra (E), Registro Info (D), Cambio Grupo Artículo (R), Cambio Descripción (W), Flete (L), Impuesto Adicional (V), Descuento (K)
- **Technologies**: React 18, Vite, SheetJS (XLSX), Lucide icons
- **Architecture**: Single-component monolith (`src/PortalSAP.jsx`) — all UI and logic in one file
- **Client-side storage**: LocalStorage for submitted requests and user's personal clusters

## Development Commands

```bash
npm install          # Install dependencies
npm run dev         # Start dev server (http://localhost:5173)
npm run build       # Build for production (output: dist/)
npm run preview     # Preview production build locally
```

## Codebase Structure

```
portal-maestros-sap/
├── src/
│   ├── PortalSAP.jsx          # Main component: entire app logic + UI (27KB)
│   └── main.jsx               # React entry point
├── datos/
│   └── maestros_muestra.js     # Embedded sample master data (534 centers, 600 SKUs, 400 suppliers)
├── index.html                 # Root HTML
├── vite.config.js             # Vite configuration
├── package.json               # Dependencies and scripts
├── README.md                  # User-facing overview
├── QUICKSTART.md              # Quick setup guide
├── NOTAS_TECNICAS.md          # Technical architecture and integration points
└── CHANGELOG.md               # Version history
```

## Architecture

### Single Component Design

The entire application is implemented as a monolithic React component in `src/PortalSAP.jsx`. This includes:
- State management (React hooks: `useState`, `useRef`, `useCallback`)
- UI layout and styling (inline CSS in `<style>` tag)
- Data validation logic
- Excel import/export (via `xlsx` library)
- LocalStorage persistence for requests and clusters

### Key Data Structures

**MAESTROS** object contains:
- `CENTROS`: map of center codes → names
- `SKUS`: array of SKU objects with `codigo`, `nombre`, `grupo`, `descripcion`
- `PROVEEDORES`: array of supplier objects with `codigo`, `nombre`

**PLANILLAS** array defines each master data module with:
- `portalId`: identifier (used by `LETRAS` map)
- `nombre`: display name
- `userCols`: columns user enters
- `autoCols`: columns auto-populated (names, descriptions)
- `salida`: output columns for SAP LSMW
- `ejemplo`: sample row
- `claveUnica` (optional): unique key fields for duplicate detection

**LETRAS** map links SAP change types (A, B, C, C1, etc.) to portal sheet IDs.

### Key Features

1. **Guía de Solicitudes**: 25 types of modifications with auto-selection of required sheets
2. **Editable Grid**: duplicate rows, import/export Excel, date masking (DD-MM-AAAA), default final date 31-12-9999
3. **Validation**: centers/SKUs/suppliers exist in master data; duplicate row detection
4. **Center Expansion**: comma-separated codes or cluster names expand to one row per center
5. **Gestor de Locales**: convert vertical list to horizontal; create/manage personal clusters
6. **Centro de Ayuda**: visual request flow and detailed catalog of each sheet
7. **Bandeja de Solicitudes**: history of submissions with states (Enviada, En proceso, Aplicada, Rechazada); download processed files

### Persistence Model

- **Solicitudes (requests)**: shared storage (all portal users see them) — intended for Data Masters team to manage
- **Clusters (personal)**: per-user storage (each user sees only theirs)

If no storage available, falls back to session memory. For production, replace with a database.

## Integration Points for Production

From `NOTAS_TECNICAS.md`, these areas require backend integration:

1. **Master Data from Database**: Replace `MAESTROS` object and `cargarMaestro` function with API calls (`GET /centros`, `/skus`, `/proveedores`, or a batch validation endpoint)
2. **Submit Requests**: `guardarSolicitud` must call backend (`POST /solicitudes`) to save folio, requester, validated rows, state, and generate/route LSMW file
3. **Authentication & Roles**: Add login and separate:
   - *Requester*: creates requests, sees only their own
   - *Data Masters Team*: sees all, changes states
4. **Uniqueness Rules per Module** (PENDING): Confirm and set `claveUnica` array for each sheet to enforce correct unique key constraints

## Key Functions and Patterns

Look for these in `PortalSAP.jsx`:

- `useState` hooks: manage active tab, planilla data, validation results, request list
- Event handlers: `handleInputChange`, `handleValidar`, `guardarSolicitud`, `handleImportarExcel`
- Validation logic: check if center/SKU/supplier exists, date format, price > 0, duplicate rows
- Excel I/O: `XLSX.utils.json_to_sheet()`, `XLSX.read()` for import
- Cluster expansion: expand comma-separated centers and cluster names to individual rows
- Inline styles: CSS-in-JS with class names like `.sol-wrap`, `.btn-validar`, `.grid-row`

## Development Notes

- The component is large (~27KB) but monolithic by design. If refactoring, preserve the state dependencies and inline styles
- Master data sample is embedded to allow immediate testing. Full data loaded via Excel button
- All validation happens client-side; dates follow DD-MM-AAAA format with automatic masking
- The `LETRAS` map connects SAP letter codes to portal sheets; maintain this mapping when adding new modules
- LocalStorage keys used: check for serialization/deserialization patterns in the code
- Icons from Lucide React — import additional icons from the package as needed
- No TypeScript; JSX with plain React hooks
