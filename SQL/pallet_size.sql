SELECT
    stkhstm.warehouse,
    stkhstm.product,
    stockm.long_description,
    stockm.analysis_a,
    stockm.analysis_b,
    stkhstm.dated,
    stkhstm.movement_reference,
    stkhstm.transaction_type,
    stkhstm.comments,
    ABS(stkhstm.movement_quantity) AS movement_quantity
FROM
    ibco.scheme.stkhstm stkhstm
JOIN
    ibco.scheme.stockm stockm
ON
    stockm.product = stkhstm.product
    AND stockm.warehouse = stkhstm.warehouse
WHERE
    stkhstm.dated >= GetDate() - 720
    AND (
        (stkhstm.transaction_type = 'BINT' AND stkhstm.comments = 'From warehouse 01')
        OR
        (stkhstm.transaction_type = 'TRAN' AND stkhstm.comments = 'W/house transfers from 01 to 00')
    );
