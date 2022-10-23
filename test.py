solar_totals = [14272, 12952]

low = [solar_totals[0]*0.024, solar_totals[1]*0.024]
high = [solar_totals[0]*0.02218, solar_totals[1]*0.02218]
low = [solar_totals[0] - low[0], solar_totals[1] - low[1]]
high = [solar_totals[0] + high[0], solar_totals[1] + high[1]]

ranges_low = []
for i in low:
    i = '{:0,.0f}'.format(i)
    ranges_low.append(i)

ranges_high = []
for i in high:
    i = '{:0,.0f}'.format(i)
    ranges_high.append(i)

print(ranges_low)
print(ranges_high)