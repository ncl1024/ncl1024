'声明两个变量
Dim x,y
'赋值给用户输入
x = inputbox("Number1...")
y = inputbox("Number2...")
'声明五个变量
Dim z,n,m,a,b
'关系运算符
z = x > y
n = x < y
m = x <> y
a = x >= y
b = x <= y
'输出运算结果 
msgbox("Da yu:" & z)
msgbox("Xiao yu:" & n)
msgbox("Bu deng yu:" & m)
msgbox("Da yu deng yu:" & a)
msgbox("Xiao yu deng yu:" & b)
'声明三个变量
Dim c,d,e
'逻辑运算符
    '或运算 只要有一个为True 结果为True
c = (z or n or  m or a or b)
    '或运算 只要有一个为False 结果为False
d = (z and n and m and a and b )
    '取反操作 d,为True 结果为False
e = (not d)
'输出运算结果
msgbox("Huo:" & c)
msgbox("Yu:" & d)
msgbox("Fei:" & e)
