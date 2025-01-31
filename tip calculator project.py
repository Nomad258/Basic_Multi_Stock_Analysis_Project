total_bill = float(input ("what is the total bill?"))
tip_percentage = float(input ("what percentage tip would you like to give?"))
tip = total_bill * tip_percentage / 100
print(f"Your tip is ${tip}")