name = "mahmoud"
age= "18"
print(f"my name is {name} and my age is {age}")
for i in range(1,11):
    print(f"{i} x 5 = {i*5}")
numbers =[1,2,3,4,5,6,7,8,9,10]
def is_even(x):
    if x%2==0:
        return True
    return False
for n in numbers:
    print(f"{n} is even: {is_even(n)}")