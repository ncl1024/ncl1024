Dim s,i
i = 1
s = 0
for i = 1 to 100 step 1
     
    if i mod 3 = 0 then
    s = s + 1
    end if
next
msgbox(s)