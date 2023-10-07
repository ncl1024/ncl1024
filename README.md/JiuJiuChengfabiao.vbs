    '声明两个变量
    Dim i,j
    '初始赋值
    i = 1
    j = 1
    '按不同的条件进行for循环
    For i = 1 to 9 step 1
        For j = 1 to 9 step 1
       str = str & i & " x " & j & " = " & i * j & vbEnter & "  "
        next 
    next  
    msgbox(str)