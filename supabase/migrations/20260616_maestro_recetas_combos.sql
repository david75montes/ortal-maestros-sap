-- Tabla maestra de recetas
CREATE TABLE IF NOT EXISTS maestro_recetas (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  codigo_sap text UNIQUE NOT NULL,
  nombre text NOT NULL,
  insumos jsonb NOT NULL DEFAULT '[]',
  organizacion_id text,
  updated_at timestamptz DEFAULT now()
);

-- Tabla maestra de combos
CREATE TABLE IF NOT EXISTS maestro_combos (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  codigo_sap text UNIQUE NOT NULL,
  nombre text NOT NULL,
  pasos jsonb NOT NULL DEFAULT '[]',
  organizacion_id text,
  updated_at timestamptz DEFAULT now()
);

-- RLS: sólo usuarios autenticados pueden leer y escribir
ALTER TABLE maestro_recetas ENABLE ROW LEVEL SECURITY;
ALTER TABLE maestro_combos  ENABLE ROW LEVEL SECURITY;

CREATE POLICY "auth_read_recetas"  ON maestro_recetas FOR SELECT TO authenticated USING (true);
CREATE POLICY "auth_write_recetas" ON maestro_recetas FOR ALL    TO authenticated USING (true) WITH CHECK (true);

CREATE POLICY "auth_read_combos"   ON maestro_combos  FOR SELECT TO authenticated USING (true);
CREATE POLICY "auth_write_combos"  ON maestro_combos  FOR ALL    TO authenticated USING (true) WITH CHECK (true);
