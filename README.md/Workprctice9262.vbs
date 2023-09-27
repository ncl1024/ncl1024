'声明一个变量
Dim num
'赋值输入
num = inputbox("Please enter One Num...")
'转换数值类型
num = Cint(num)
'根据不同的条件做判断
if num > 0 then
   msgbox("Zheng Shu!")
elseif num = 0 then
   msgbox("Deng Yu 0!")
elseif num < 0 then
   msgbox("Fu Shu!")
else 
   msgbox("Qing Chong Xin Shu Ru!")
end if