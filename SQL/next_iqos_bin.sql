WITH CommonQuantities AS (
    SELECT
        stkhstm.product,
        ABS(stkhstm.movement_quantity) AS movement_quantity,
        COUNT(*) AS frequency,
        RANK() OVER (PARTITION BY stkhstm.product ORDER BY COUNT(*) DESC) AS rank
    FROM
        ibco.scheme.stkhstm stkhstm
    WHERE
        stkhstm.dated >= GETDATE() - 720
        AND (
            (stkhstm.transaction_type = 'BINT' AND stkhstm.comments = 'From warehouse 01')
            OR
            (stkhstm.transaction_type = 'TRAN' AND stkhstm.comments = 'W/house transfers from 01 to 00')
        )
    GROUP BY
        stkhstm.product, ABS(stkhstm.movement_quantity)
),
FilteredCommonQuantities AS (
    SELECT
        product,
        movement_quantity AS common_quantity
    FROM
        CommonQuantities
    WHERE
        rank = 1
),
FilteredBins AS (
    SELECT
        stquem.warehouse,
        stquem.product,
        stquem.sequence_number, -- Include sequence number for sorting
        stquem.bin_number,      -- Include bin number in the output
        stquem.quantity_free,
        stquem.best_before_date,
        cq.common_quantity,
        SUM(stquem.quantity_free) OVER (
            PARTITION BY stquem.product
            ORDER BY
                stquem.best_before_date ASC,
                stquem.sequence_number ASC -- Sort by sequence number after best-before date
            ROWS UNBOUNDED PRECEDING
        ) AS cumulative_quantity
    FROM
        [ibco].[scheme].[stquem] AS stquem
    LEFT JOIN
        FilteredCommonQuantities AS cq
    ON
        stquem.product = cq.product
    WHERE
        stquem.warehouse NOT IN ('00', '98', '03', '02') -- Exclude specific warehouses
        AND stquem.quantity_free > 0                    -- Only include bins with stock
        AND stquem.best_before_date > GETDATE() + 40    -- Only include bins with valid expiration dates
)
SELECT
    warehouse,
    product,
    sequence_number,           -- Include the sequence number in output
    bin_number,                -- Include the bin number in output
    quantity_free,
    best_before_date,
    common_quantity,
    cumulative_quantity
FROM
    FilteredBins
WHERE
    cumulative_quantity <= common_quantity + 10 -- Allow slight overage in total quantity
ORDER BY
    best_before_date ASC,  -- Sort primarily by earliest expiration date
    sequence_number ASC;   -- Sort secondarily by sequence number
