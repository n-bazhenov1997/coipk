from math import factorial

n = int(input())
result = [factorial(n) // (factorial(k) * factorial(n - k)) for k in range(n + 1)]
print(result)
