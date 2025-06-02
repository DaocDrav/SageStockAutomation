 SELECT  "opheadm"."order_no", "opdetm"."product", "opheadm"."status", "opdetm"."order_line_no", "opheadm"."date_entered",  "stockm"."analysis_b", "stockm"."description", "stockm"."bin_number", "stockm"."long_description", "oppickm"."bin", "oppickm"."warehouse", "oppickm"."quantity", "stockm"."physical_qty", "opheadm"."customer_order_no", "stockm"."allocated_qty"
 FROM   ("ibco"."scheme"."oppickm" "oppickm" INNER JOIN ("ibco"."scheme"."opheadm" "opheadm" INNER JOIN "ibco"."scheme"."opdetm" "opdetm" ON "opheadm"."order_no"="opdetm"."order_no") ON ("oppickm"."order_no"="opdetm"."order_no") AND ("oppickm"."component"="opdetm"."product")) INNER JOIN "ibco"."scheme"."stockm" "stockm" ON "opdetm"."product"="stockm"."product"
 WHERE  "opheadm"."status"<'7' AND "opheadm"."customer"='00' AND "stockm"."warehouse"='00'
 ORDER BY "opdetm"."order_no"


