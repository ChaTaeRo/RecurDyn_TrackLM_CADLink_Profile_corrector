import csv

f = open("E:\\test\\fileFromExcel.csv", 'r', encoding='utf-8')
rdr = csv.reader(f, quoting=csv.QUOTE_NONNUMERIC)

results = []
for row in rdr:
    results.append(row)
f.close()

nRow = len(results)
for i in range(nRow):
    sum = results[i][0]+results[i][1]
    distance = results[i][0]*results[i][0]+results[i][1]*results[i][1]
    results[i].append(distance)
    results[i].insert(0,i+1)

min1 = 0
min2 = 0
min1_index = 0
min2_index = 0

for i in range(nRow):
    num = results[i][3]
    if i == 0:
        min1 = num
        min2 = num
    if num < min1:
        min1 = num
        min1_index = i

for i in range(nRow):
    num = results[i][3]
    if min2 > num > min1:
        min2 = num
        min2_index = i

start_number = 0
second_start_number = 0

if results[min1_index][1] > 0:
    start_number = min1_index
else:
    start_number = min2_index

second_start_number = nRow - start_number
new_order = []
for i in range(start_number+1):
    new_order.append(start_number)
    start_number = start_number - 1

for i in range(second_start_number-1):
    new_order.append(nRow-i-1)

new_profile = []

new_profile.append([0,0,0.001,0])
for num in new_order:
    new_profile.append(results[num])
new_profile.append([0,0,0.001,0])

f = open("E:\\test\\changedProfile.csv",'w')
for i in range(len(new_profile)):
    # i = i - 1
    f.write(str(new_profile[i][1])+',' + str(new_profile[i][2]) + '\n')
f.close()

print("Mission completed")