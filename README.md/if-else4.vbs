'声明两个变量
Dim Age
'赋值给用户输入
Age = inputbox("Number1...")
'根据不同的条件做判断
if Age >= 0 and Age <= 10 then
     msgbox("Teenagers!")
elseif Age > 12 and Age <= 25 then
     msgbox("Young!")
elseif Age > 25 and Age <= 40 then
     msgbox("Zhuang Nian!")
elseif Age > 40 and Age < 60 then
     msgbox("Zhong Nian!")
elseif Age >= 60 then 
     msgbox("Old Age!")
end if