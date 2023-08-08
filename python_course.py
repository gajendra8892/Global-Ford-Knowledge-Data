# yes both numbers and new_numbers we give us the same output

number = float(input("enter the number:"))

if number > 0:
    print("the number is positive")
elif number == 0:
    print("the number is zero")
else:
    print("the number is negative")

hrs = input("Enter Hours:")
h = float(hrs)
hourly_wage = input("Enter the Rate:")
x = float(hourly_wage)
if h > 40:
	print(40* x + (h-40)*1.1*x)

D1 = {'list1': [4, 7, 10, 20], 'list2': [7, 16, 9, 10], 'list3': [13, 10, 4, 8], 'list4': [7, 20, 6, 11]}

combined_values = []
for values in D1.values():
    combined_values.extend(values)

values_without_dup = set(combined_values)

output = list(values_without_dup)
print(output)

def count_repeated_num(input_list):
    repeated_dict = {}
    for item in input_list:
        if item in repeated_dict:
            repeated_dict[item] += 1
        else:
            repeated_dict[item] = 1
    return repeated_dict

List1 = [1, 2, 2, 3, 4, 1, 4, 5, 5, 6, 7, 7]
output1 = count_repeated_num(List1)
print(output1)

sample_data = [(), (), ('',), ('a', 'b'), ('a', 'b', 'c'), ('d')]

output2 = []
for tup in sample_data:
    if tup != ():
        output.append(tup)

print(output2)



def get_float_value(tuple_item):
    return float(tuple_item[1])

data = [('item1', '12.20'), ('item2', '15.10'), ('item3', '24.5')]

sorted_data = sorted(data, key=get_float_value, reverse=True)

print(sorted_data)
