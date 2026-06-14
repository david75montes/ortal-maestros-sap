import { createClient } from "@supabase/supabase-js";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

// Leer .env manualmente (sin dotenv)
const __dir = dirname(fileURLToPath(import.meta.url));
const envPath = join(__dir, "../.env");
const env = Object.fromEntries(
  readFileSync(envPath, "utf8")
    .split("\n")
    .filter(l => l.includes("="))
    .map(l => l.split("=").map(s => s.trim()))
);

const supabase = createClient(env.VITE_SUPABASE_URL, env.VITE_SUPABASE_SERVICE_KEY);

// Importar datos de muestra
const { CENTROS, SKUS, PROVEEDORES } = await import("../datos/maestros_muestra.js");

async function cargarTabla(tabla, filas, label) {
  console.log(`\n→ Cargando ${label} (${filas.length} registros)...`);
  const BATCH = 500;
  let ok = 0;
  for (let i = 0; i < filas.length; i += BATCH) {
    const lote = filas.slice(i, i + BATCH);
    const { error } = await supabase.from(tabla).upsert(lote, { onConflict: "codigo" });
    if (error) {
      console.error(`  ✗ Error en lote ${i}-${i + BATCH}:`, error.message);
    } else {
      ok += lote.length;
      process.stdout.write(`  ✓ ${ok}/${filas.length}\r`);
    }
  }
  console.log(`  ✓ ${label}: ${ok} registros cargados.`);
}

// Convertir objetos {codigo: nombre} a arrays de filas
const centros    = Object.entries(CENTROS).map(([codigo, nombre]) => ({ codigo, nombre }));
const skus       = Object.entries(SKUS).map(([codigo, nombre]) => ({ codigo, nombre }));
const proveedores = Object.entries(PROVEEDORES).map(([codigo, nombre]) => ({ codigo, nombre }));

await cargarTabla("centros",    centros,    "Centros");
await cargarTabla("skus",       skus,       "SKUs");
await cargarTabla("proveedores", proveedores, "Proveedores");

console.log("\n✅ Carga completa.");
