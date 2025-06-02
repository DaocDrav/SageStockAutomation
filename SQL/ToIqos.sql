SELECT
    stockm.alpha,
    stockm.product,
    stockm.long_description,
    stockm.analysis_a,
    stockm.analysis_b,
    stockm.analysis_c,
    stockm.min_stock_level AS 'MinLevel',
    stockm.maximum_stock_leve AS 'MaxLevel',
    (stockm.maximum_stock_leve - stockm.physical_qty) AS 'SuggestedQty',
    Pinn_00FreeStock.FreeStock AS '00FreeStk',
    Pinn_01FreeStock.FreeStock AS '01FreeStk',
    Pinn_96FreeStock.FreeStock AS '96FreeStk',
    stockm.bin_number AS '00Bin',
    MAX(stallocm.order_no) AS 'OrderNo',
    SUM(stallocm.qty_allocated) AS 'TotalAllocated',
    MAX(oppickm.bin) AS 'OppickBin'
FROM
    ibco.scheme.stockm stockm
    JOIN ibco.dbo.Pinn_00FreeStock Pinn_00FreeStock ON Pinn_00FreeStock.product = stockm.product
    JOIN ibco.dbo.Pinn_01FreeStock Pinn_01FreeStock ON Pinn_01FreeStock.product = stockm.product
    JOIN ibco.dbo.Pinn_96FreeStock Pinn_96FreeStock ON Pinn_96FreeStock.product = stockm.product
    LEFT JOIN ibco.scheme.stallocm stallocm ON stallocm.product = stockm.product
    LEFT JOIN ibco.scheme.oppickm oppickm ON oppickm.order_no = stallocm.order_no
                                         AND stallocm.product = oppickm.component
                                         AND stallocm.warehouse = oppickm.warehouse
                                         AND stallocm.qty_allocated = oppickm.quantity
                                         AND stallocm.order_line = oppickm.order_detail
WHERE
    stockm.warehouse = '00'
    AND Pinn_00FreeStock.FreeStock < stockm.min_stock_level
    AND (
        stallocm.qty_allocated IS NULL
        OR stallocm.qty_allocated = 0
        OR (
            stallocm.qty_allocated > 0
        )
    )
    AND (Pinn_01FreeStock.FreeStock + Pinn_96FreeStock.FreeStock > 0)
    AND NOT EXISTS (
        SELECT 1
        FROM ibco.scheme.oppickm opp
        INNER JOIN ibco.scheme.opheadm head ON opp.order_no = head.order_no
        INNER JOIN ibco.scheme.opdetm det ON head.order_no = det.order_no
        WHERE det.product = stockm.product
          AND head.status < '7'
          AND head.customer = '00'
          AND stockm.warehouse = '00'
    )
GROUP BY
    stockm.alpha,
    stockm.product,
    stockm.long_description,
    stockm.analysis_a,
    stockm.analysis_b,
    stockm.analysis_c,
    stockm.min_stock_level,
    stockm.maximum_stock_leve,
    stockm.physical_qty,
    Pinn_00FreeStock.FreeStock,
    Pinn_01FreeStock.FreeStock,
    Pinn_96FreeStock.FreeStock,
    stockm.bin_number
ORDER BY
    stockm.alpha;
