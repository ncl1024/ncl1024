'声明一个变量
Dim r 
'将变量进行赋值
r = 1
'判断
do 
  '当r中存储的数字值小于等于50，这些数值都将要被5进行整除判断
  if r mod 5 = 0 then
  '输出r
  msgbox(r)
  end if
  '每判断完一个数，接下来获取下一个数字
  r = r + 1
loop while r <= 50
