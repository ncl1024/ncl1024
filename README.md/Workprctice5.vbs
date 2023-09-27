'声明三个变量
Dim R,y,K,T
'输入赋值
R = inputbox("Please into Write...")
y = inputbox("Please into Write...")
'定义两个函数的数据类型
R = Cint(R)
y = Cint(y)
'T = 长 * 2  K = 宽 * 2
T = 2 * R
K = 2 * y
'声明两个变量
Dim m,n
'm = 周长    n = 面积
n = R * y
m = T + K
'输出
msgbox(n)
msgbox(m)