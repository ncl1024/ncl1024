'声明两个变量
Dim x,y
'初始化
x = 1
y = 0
'累加计算
do
  '只要x中得值小于等于5,就要做累加操作
  '每一个小于等于5的数字,都要累加在y变量中作求和操作
  y = y + x
  '每加完一个数字后，还要找到另一个数字
  x = x + 1
loop while x <= 5
'只有循环完成的情况下，才代表着所有数据累加完成
msgbox("The cumulative sum between 1-5 is:" & y)