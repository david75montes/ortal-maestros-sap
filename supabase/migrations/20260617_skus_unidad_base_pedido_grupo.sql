ALTER TABLE skus
  ADD COLUMN IF NOT EXISTS unidad_base    text,
  ADD COLUMN IF NOT EXISTS unidad_pedido  text,
  ADD COLUMN IF NOT EXISTS grupo_articulo text;
