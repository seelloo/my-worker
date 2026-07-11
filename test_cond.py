import sys
sys.path.append('.')
from scripts.excel_mark_by_conditions import parse_condition, match_value

c = parse_condition('X:X="755941714"')
print('Parsed:', c)
val_in_excel = 755941714
print('Match result:', match_value(val_in_excel, c[1], c[2]))
