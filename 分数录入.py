import xlwings as xw
a='''
42. 王舒仪 80
43. 叶双源 127
44. 张笑语 118
45. 韩馥茹  98
46. 王思思 91
'''

print(a )
li=a.split()
print(li)
names=[]
c=[]
t=[]
for j in range(0,len(li)):
    if (j+1)%3==2:
        names.append(li[j])
    elif (j+1)%3==0 :
        c.append(li[j])
print(names)
print(c)
E=xw.Book(r"F:\6班月考成绩单.xlsx")
sh=E.sheets["sheet1"]

for i in range(3,61):
    n=sh.range("B"+str(i)).value
    print(n)
    for k in range(0,len(names)):
        
        if names[k]==str(n):
            print(1)
            sh.range('D'+str(i)).value=c[k]
            print(c[k])
            break
E.save(r"F:\tre.xlsx")        
