'求水仙花数
'声明一个循环变量
Dim n1,n2,n3 
For num = 100 To 999 step 1
    '分解数字的各位
    n1 = Int(num / 100)
    n2 = Int((num Mod 100) / 10)
    n3 = num Mod 10
    '定义一个变量
    Dim c
    ' 计算立方和
    c = (n1 ^ 3) + (n2 ^ 3) + (n3 ^ 3)
    
    ' 检查是否是水仙花数
    If c = num Then
        msgbox(num)
    End If
Next
