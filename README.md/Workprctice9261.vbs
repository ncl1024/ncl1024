'声明一个变量
Dim Month
'赋值输入
Month = inputbox("Please enter your score!")
'转换数值类型
Month = Cint(Month)
'根据不同的条件做判断
select case Month
case 1,2,3  msgbox("Spring")
case 4,5,6  msgbox("Summer")
case 7,8,9  msgbox("Autumn")
case 10,11,12  msgbox("Winter")
case else msgbox("Already served!")
end select