'统计1-100之间能够被3整除的数据的总个数 for
'声明两个变量
Dim s,i
'初始化
i = 1
s = 0

for i = 1 to 100 step 1
     
    if i mod 3 = 0 then
    s = s + 1
    end if
next
msgbox(s)