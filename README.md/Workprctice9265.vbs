'声明一个变量
Dim Week
'输入赋值
Week = inputbox("Please Enter...")
'将赋值的数转化为整型
Week = Cint(Week)
'对照不同的条件做出回应
select case Week
case 1 msgbox("Monday")
case 2 msgbox("Tuesday")
case 3 msgbox("Wednesday")
case 4 msgbox("Thursday")
case 5 msgbox("Friday")
case 6 msgbox("Saturday")
case 7 msgbox("Sunday")
case else msgbox("Please Enter Again!")
end select