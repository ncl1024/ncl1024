'声明一个变量
Dim Year
'输入赋值
Year = inputbox("Please Enter...")
'将赋的值转化成整型
Year = Cint(Year)
'根据不同的条件做出判断
if (Year Mod 4 = 0 And Year Mod 100 <> 0) Or (Year Mod 400 = 0)  Then
   msgbox("Run Nian!")
else 
   msgbox("Bu Shi Run Nian!")
end if