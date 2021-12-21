def sum(num1, num2):
    sum1 = num1 + num2
    return sum1
    
#this is where the function values get printed
print(sum(4, 5))

 #this is where the function values get printed

def subtract(num1, num2):
    if num1 >= num2:
        difference = num1 - num2
    else:
        difference = f"{num1} is smaller than {num2}"
    return difference
 #this is where the function subtract values get printed

print(subtract(9,3))
print(subtract(8,9))