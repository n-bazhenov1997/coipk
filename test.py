def dec_to_hex(dec):
    result = ''
    hex_nums = [
        '0', '1', '2', '3', 
        '4', '5', '6', '7', 
        '8', '9', 'A', 'B', 
        'C', 'D', 'E', 'F'
    ]
    while dec != 0:
        result += hex_nums[dec % 16]
        dec //= 16
    result = result[::-1].rjust(2, '0')
        
    return result

def rgb(r, g, b):
    result = ''
    decs_nums = (r, g, b)
    for n in decs_nums:
        result += dec_to_hex(n)
        
    return result

print(rgb(-20, 275, 125))