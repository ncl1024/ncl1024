'声明两个变量
'Dim x,y
'赋值给用户输入
'x = inputbox("Number1...")
'y = inputbox("Number2...")
'声明五个变量
'Dim z,n,m
'z = Cint(z)
'n = Cint(n)
'm = Cint(m)
'判断
'if x > y or x < y or x = y then
   'msgbox("Da yu!")
  ' msgbox("Xiao yu!")
   'msgbox("Deng yu!")
'end if 





'根据用户输入的两个数字，判断这两个数字，是大于关系，还是小于关系，或者是等于关系
'格式一 if语句
'声明一个变量，用来接收输入的数据
Dim xy
'赋值 X=47 y=89
x=inputbox("number01:")
y=inputbox("numbero2:")
'三个关系的判断，分别使用三个if语句来进行判断操作
'if语句 大于关系 47>89 False
if x>y then
   msgbox("x da yu y")
end if
'if语句 小于关系 47<89 True
if x<y then
   msgbox("x xiao yu y")
end if
'if语句 等于关系 47=89 False
if x=y then
   msgbox("x deng yu y")
end if