'声明一个变量
Dim Result
'赋值输入
Result = inputbox("Please enter your score!")
'转换数值类型
Result = Cint(Result)
'根据不同的条件做判断
if Result >= 90 then
   msgbox("A")
elseif Result >= 80 then
   msgbox("B")
elseif Result >= 70 then
   msgbox("C")
elseif Result >= 60 then
   msgbox("D")
else 
   msgbox("F")
end if