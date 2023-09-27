'声明一个变量
Dim x
'赋值
x = inputbox("Please enter a 4-digit number...")
'声明四个变量
Dim y,z,m,n
y = x mod 10
z = x \ 10 mod 10
m = x \ 100 mod 10
n = x \ 1000  

msgbox(y)

msgbox(z)

msgbox(m)

msgbox(n)