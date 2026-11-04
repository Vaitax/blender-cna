-- Listas de materiales de las mezclas MGM (equivale a la hoja "Insumos por producto")
-- Columnas: Codigos, Name, Código Insumo, ItemName, Dosificación, Unidad Medida Insumo
-- {schema} se reemplaza por la base configurada (por defecto SBO_CNA_CHILE). Solo lectura.
SELECT
  T0."Code"        AS "Codigos",
  P."ItemName"     AS "Name",
  T1."Code"        AS "CodigoInsumo",
  C."ItemName"     AS "ItemName",
  T1."Quantity"    AS "Dosificacion",
  C."InvntryUom"   AS "Unidad"
FROM "{schema}"."OITT" T0
INNER JOIN "{schema}"."ITT1" T1 ON T1."Father" = T0."Code"
INNER JOIN "{schema}"."OITM" P  ON P."ItemCode" = T0."Code"
LEFT  JOIN "{schema}"."OITM" C  ON C."ItemCode" = T1."Code"
WHERE T0."Code" LIKE 'MGM%'
  AND COALESCE(P."frozenFor", 'N') <> 'Y'   -- sin artículos marcados como Inactivo en SAP
ORDER BY T0."Code", T1."ChildNum"
