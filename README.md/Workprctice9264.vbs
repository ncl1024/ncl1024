'声明一个变量
Dim num
'输入赋值
num = inputbox("Please Enter...")
'将赋的值转化成整型
num = Cint(num)
'根据不同的条件做出判断
if num Mod 2 = 0 Then
   msgbox("Ou Shu!")
else 
   msgbox("Ji Shu!")
end if