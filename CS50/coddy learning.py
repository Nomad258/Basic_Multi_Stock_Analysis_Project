b1 = True
b2 = False
b3 = b1 or b2

print (f"b3 = {b3}")   # b3 = True  

#########

a = 6
b = 4
c = a * b
print (f"c = {c}")   # c = 24

#########`
b1 = True
b2 = True
b3 = False

# Don't change the line below
b4 = b1 and b2 and (not b3)
print(f"b4 = {b4}")     # b4 = True

#########

wind = int(input()) # Don't change this line
status = "unset"
# Type your code below
if wind < 8:
    print ("calm")
elif wind >= 8 and wind <= 31:
    print ("Breeze")
else:
    print ("Gale")

wind = 54
# Don't change the line below
print(f"status = {status}")     # status = Gale

#########

