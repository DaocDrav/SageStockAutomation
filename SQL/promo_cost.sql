UPDATE scheme.stockm
SET current_cost = p.price6
FROM dbo.Pinn_Live_00_Promotions p
WHERE stockm.product = p.product
  AND stockm.warehouse = p.warehouse
  AND p.price6 < stockm.current_cost
  AND stockm.warehouse = '00'
  AND p.date_end > GETDATE();