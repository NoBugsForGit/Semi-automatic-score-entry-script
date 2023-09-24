import xlwings as xw
a='''
name
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
E=xw.Book(r"F:\.xlsx")
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
