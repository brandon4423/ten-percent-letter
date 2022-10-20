solar_totals = [14272, 12952]
math_1 = solar_totals.copy()
math_2 = solar_totals.copy()

range1_percent = 0.024 * math_1[0]
rangemath1 = math_1[0] - range1_percent

range2_percent = 0.024 * math_1[1]
rangemath2 = math_1[1] - range2_percent

range3_percent = 0.02218 * math_2[0]
rangemath3 = math_2[0] + range3_percent

range4_percent = 0.02218 * math_2[1]
rangemath4 = math_2[1] + range4_percent

range_totals_1 = [rangemath1, rangemath2]
range_totals_2 = [rangemath3, rangemath4]

ranges_low = []
for i in range_totals_1:
    i = '{:0,.0f}'.format(i)
    ranges_low.append(i)

ranges_high = []
for i in range_totals_2:
    i = '{:0,.0f}'.format(i)
    ranges_high.append(i)

print(ranges_low)
print(ranges_high)

