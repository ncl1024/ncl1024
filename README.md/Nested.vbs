'声明三个变量
Dim x,y,z 
'输入赋值
x = inputbox("Please Enter... ")
y = inputbox("Please Enter... ")
z = inputbox("Please Enter... ")
'转换数值类型
x = Cint(x)
y = Cint(y)
z = Cint(z)
'根据不同的条件做判断
if x > y then
    msgbox("x")
elseif y > z then
    msgbox("y")
elseif z > x then
    msgbox("z")
end if

