
wb_structure = [
    ('CDTs', 'sht1_tbl1_fmls', 'A4:N13'),
    ('CDTs', 'sht1_tbl1_vals', 'A4:N13')
]

sht1_tbl1_fmls = """Cuenta	Emisor	Id	Capital	Fecha Constitución	Fecha Vencimiento	Plazo	Fecha último Pago	Interés Efectivo	Días hasta el 31/12/23	Interés Acumulado	Cotización Bolsa (a)	Inc. Cotización Bolsa	Incremento fiscal
													
122505025	COLTEFINANCIERA	1577316	20000000	45187	45369	=+DAYS360(E6;F6)	45278	0.15	=+DAYS360(H6;DATE(2023;12;31))	=+((1+I6)^(J6/360) - 1)*D6	1.004	=+(L6-1)*D6	=+M6+K6
122505026	COLTEFINANCIERA	1579122	20000000	45287	45470	=+DAYS360(E7;F7)	45287	0.148	=+DAYS360(H7;DATE(2023;12;31))	=+((1+I7)^(J7/360) - 1)*D7	1.004	=+(L7-1)*D7	=+M7+K7
122505027	COLTEFINANCIERA	1579121	10000000	45287	45470	=+DAYS360(E8;F8)	45287	0.148	=+DAYS360(H8;DATE(2023;12;31))	=+((1+I8)^(J8/360) - 1)*D8	1.004	=+(L8-1)*D8	=+M8+K8
122505028	COLTEFINANCIERA	1579146	10000000	45287	45653	=+DAYS360(E9;F9)	45287	0.149	=+DAYS360(H9;DATE(2023;12;31))	=+((1+I9)^(J9/360) - 1)*D9	1.004	=+(L9-1)*D9	=+M9+K9
122505029	COLTEFINANCIERA	1579147	20000000	45287	45653	=+DAYS360(E10;F10)	45287	0.149	=+DAYS360(H10;DATE(2023;12;31))	=+((1+I10)^(J10/360) - 1)*D10	1.004	=+(L10-1)*D10	=+M10+K10
122505030	COLTEFINANCIERA	1579148	7993390	45287	45653	=+DAYS360(E11;F11)	45287	0.149	=+DAYS360(H11;DATE(2023;12;31))	=+((1+I11)^(J11/360) - 1)*D11	1.004	=+(L11-1)*D11	=+M11+K11
122505021	PICHINCHA	1048628	10000000	45288	45471	=+DAYS360(E12;F12)	45288	0.153	=+DAYS360(H12;DATE(2023;12;31))	=+((1+I12)^(J12/360) - 1)*D12	1	=+(L12-1)*D12	=+M12+K12
		TOTAL	=+SUM(D6:D12)							=+SUM(K6:K12)		=+SUM(M6:M12)	=+SUM(N6:N12)
"""

sht1_tbl1_vals = """Cuenta	Emisor	Id	Capital	Fecha Constitución	Fecha Vencimiento	Plazo	Fecha último Pago	Interés Efectivo	Días hasta el 31/12/23	Interés Acumulado	Cotización Bolsa (a)	Inc. Cotización Bolsa	Incremento fiscal
													
122505025	COLTEFINANCIERA	1577316	20_000_000.00	18/09/23	18/03/24	180	18/12/23	15.0000%	13	101_194.33	100.40%	80_000.00	181_194.33
122505026	COLTEFINANCIERA	1579122	20_000_000.00	27/12/23	27/06/24	180	27/12/23	14.8000%	4	30_694.93	100.40%	80_000.00	110_694.93
122505027	COLTEFINANCIERA	1579121	10_000_000.00	27/12/23	27/06/24	180	27/12/23	14.8000%	4	15_347.46	100.40%	40_000.00	55_347.46
122505028	COLTEFINANCIERA	1579146	10_000_000.00	27/12/23	27/12/24	360	27/12/23	14.9000%	4	15_444.36	100.40%	40_000.00	55_444.36
122505029	COLTEFINANCIERA	1579147	20_000_000.00	27/12/23	27/12/24	360	27/12/23	14.9000%	4	30_888.72	100.40%	80_000.00	110_888.72
122505030	COLTEFINANCIERA	1579148	7_993_390.00	27/12/23	27/12/24	360	27/12/23	14.9000%	4	12_345.28	100.40%	31_973.56	44_318.84
122505021	PICHINCHA	1048628	10_000_000.00	28/12/23	28/06/24	180	28/12/23	15.3000%	3	11_870.98	100.00%	0.00	11_870.98
		TOTAL	97_993_390.00							217_786.05		351_973.56	569_759.61
"""

