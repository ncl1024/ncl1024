'模拟一个点菜系统
'声明一个变量
Dim x 
'输入赋值
x = inputbox("Please enter the dish name...")

select case x 
case 1  msgbox("Fried chicken!")
case 2  msgbox("Take out the egg tarts!")
case 3  msgbox("Chicken nuggets cooked outside the palace!")
case 4  msgbox("Boiled Fish!")
case 5  msgbox("Sauce elbow!")
case else msgbox("Already served!")
end select