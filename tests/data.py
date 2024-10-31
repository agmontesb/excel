import sys
import os
sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


sht1_tbl1_vals = """Tabla 1:			
uno	10	55	65
dos	20	75	95
tres	30	95	125
cuatro	40	115	2,875
cinco	50	135	185
Total	150	475	472,875
""".replace(',', '.')

sht1_tbl2_vals = """Tabla 2:		
l1	300	6000
l2	150	3000
toal	450	9000
"""

sht2_tbl1_vals = """tabla 1:			
l1	700	250	950
l2	1200	230	1430
l3	340	100	440
l4	128	85	213
l5	45	18	63
	2413	683	3096
"""

sht2_tbl2_vals = """tabla 2:			
l1	25	100	0
l2	10	25	35
l3	15	38	53
l4	80	2438	2518
	105	2501	2606
"""

sht3_tbl1_vals = """l1	10	250	240
l2	0	230	230
l3	30	100	70
l4	40	85	45
l5	50	18	-32
	130	683	553
"""

sht1_tbl1_fmls = """Tabla 1:			
uno	10	55	=+G4+H4
dos	20	75	=+G5+H5
tres	30	95	=+G6+H6
cuatro	40	115	=+H7/G7
cinco	50	135	=+G8+H8
Total	=+SUM(G4:G8)	=+SUM(H4:H8)	=+SUM(I4:I8)
"""

sht1_tbl2_fmls = """Tabla 2:		
l1	300	=20*G13
l2	150	=20*G14
toal	=+SUM(G13:G14)	=+SUM(H13:H14)
"""

sht2_tbl1_fmls = """tabla 1:			
l1	700	250	=+G4+F4
l2	1200	230	=+G5+F5
l3	340	100	=+G6+F6
l4	128	85	=+G7+F7
l5	45	18	=+G8+F8
	=+SUM(F4:F8)	=+SUM(G4:G8)	=+SUM(H4:H8)
"""

sht2_tbl2_fmls = """tabla 2:			
l1	25	=+G6	=+$G$2*F13
l2	10	25	=+SUM(F14:G14)
l3	15	38	=+SUM(F15:G15)
l4	80	=+F9 + F13	=+SUM(F16:G16)
	=+F16+F15+F14	=+G16+G15+G14	=+H16+H15+H14
"""

sht3_tbl1_fmls = """l1	=+'No links, No parameters'!G4	=+'Parameters and inner links'!G4	=+G3-F3
l2	=+'No links, No parameters'!G5*'Parameters and inner links'!G2	=+'Parameters and inner links'!G5	=+G4-F4
l3	=+'No links, No parameters'!G6	=+'Parameters and inner links'!G6	=+G5-F5
l4	=+'No links, No parameters'!G7	=+'Parameters and inner links'!G7	=+G6-F6
l5	=+'No links, No parameters'!G8	=+'Parameters and inner links'!G8	=+G7-F7
	=+SUM(F3:F7)	=+SUM(G3:G7)	=+SUM(H3:H7)
"""


if __name__ == '__main__':
    # from utilities import TableComparator

    # test = test_workbook()

    pass