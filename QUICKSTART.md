# Inicio Rápido

## En 2 minutos

### 1. Instala Node.js
Descarga de https://nodejs.org (versión LTS). Instala y listo.

### 2. Prepara la carpeta
- Descomprime `portal-maestros-sap.zip`
- Abre VS Code
- File → Open Folder → selecciona `portal-maestros-sap`

### 3. Instala dependencias
Abre terminal en VS Code (Ctrl + `):
```bash
npm install
```
Espera ~30 segundos.

### 4. Inicia el desarrollo
```bash
npm run dev
```

Verás:
```
Local:   http://localhost:5173/
```

Copia la URL y ábrela en el navegador. **Listo, el portal está corriendo.**

### 5. Edita el código
Abre `src/PortalSAP.jsx`, haz cambios, guarda (Ctrl + S), recarga el navegador.

---

## Estructura de carpetas

```
src/PortalSAP.jsx          ← TODO el código (UI + lógica)
datos/maestros_muestra.js  ← Centros, SKU, proveedores de demo
package.json               ← Dependencias
index.html                 ← HTML raíz
```

---

## Comandos principales

```bash
npm run dev      # Desarrollo (http://localhost:5173)
npm run build    # Compilar para producción
npm run preview  # Ver compilado localmente
```

---

## Primeros cambios (prueba)

### Cambiar el título
Busca (Ctrl + F) `"Portal de Solicitudes"` en `src/PortalSAP.jsx`, reemplaza, guarda.

### Agregar un ícono
En la línea `import { ... } from "lucide-react"`, agrega el nombre del ícono (ej: `Trash2`).
Luego úsalo: `<Trash2 size={18} />`.

### Cambiar colores
Busca (Ctrl + F) `".sol-wrap"` o `".btn-validar"` en la sección `<style>`, edita `color:` o `background:`.

---

## Siguientes pasos

- Lee **NOTAS_TECNICAS.md** para entender la arquitectura
- Lee **CHANGELOG.md** para ver qué se agregó en cada versión
- Ve al **Centro de Ayuda** (pestaña en el portal) para entender cada planilla
- Cuando esté listo, sigue **README.md** para deployar a Vercel

---

## ¿Necesitas ayuda?

- **"npm: comando no encontrado"** → Reinicia VS Code después de instalar Node.js
- **"Cannot find module"** → Corre `npm install` nuevamente
- **Los cambios no se ven** → Recarga la página del navegador (F5)

¡Que disfrutes desarrollando!
