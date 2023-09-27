'声明一个变量
Dim num1
'输入赋值
num1 = inputbox("Please Enter...")
'将赋的值转换为整型类型
num1 = Cint(num1)
'根据不同的条件做回应
select case num1 Mod 2
case 0 msgbox("Even")
case 1 msgbox("Odd number") 
case else msgbox("Please enter a legal number!")
end select