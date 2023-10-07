'声明两个变量
Dim sno,ago
'创建两个数组并赋值
sno = array("s101", "s102", "s103", "s104", "s105", "s106")
ago = array(18,23,26,20,21,30,27)
'声明一个变量i,j 从0进行循环
max_age = ago(0)
for i = 0 to 5 step 1
    If ago(i) > max_age Then
        max_age = ago(i)
    End If
next
msgbox(max_age)

'声明一个变量j数组下标0开始
j = 0

'声明一个变量x,用来接收最大值
x = ago(0)
while j <= 6
   If ago(j) < x Then
        x = ago(j)
    End If
   j = j + 1
wend
msgbox(x)
'msgbox(ago(0))
'msgbox(ago(6))
'msgbox(ago(9))
