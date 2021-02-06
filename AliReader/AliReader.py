from bs4 import BeautifulSoup
import xlsxwriter


workbook = xlsxwriter.Workbook('aliexpress.xlsx')
worksheet = workbook.add_worksheet()

orders = []

fileN = 14

def readDataHTML():
    global days
    global weekdayBuckets
    global mptc
    global tptc
    global targetdir
    global fileN
    

    for i in range(1,fileN + 1):
        f = open("ali" + str(i) + ".html", "r", encoding='utf8')

        if f.mode == 'r':
            soup = BeautifulSoup(f.read(), "html.parser")

            #print("parsed",(i))

            content = soup.find_all('tbody', attrs={"class": "order-item-wraper"})
            tmp = []
            #print("read",(i),'len=',len(content))

            for order in content:
                info = order.find_all('span', attrs={'class': 'info-body'})
                id = info[0].contents[0]
                date = info[1].contents[0]
                #price = float(order.find('p', attrs={'class': 'amount-num'}).contents[0].split([' ','\n'])[0].replace(',','.'))
                orderprice = int([x for x in order.find('p', attrs={'class': 'amount-num'}).contents[0].translate({ord('\n'): None}).split(' ') if x != ' ' and x != ''][0].replace(',',''))
                items = order.find_all('tr', attrs={'class': 'order-body'})
                productsprice = 0
                products = []

                #if id == '8012802117885974':
                #    print('hr')

                exists = False
                for o in orders:
                    if o[0] == id:
                        exists = True
                if exists:
                    continue

                for p in items:
                    name = p.find('a', attrs={'class': 'baobei-name'}).contents[0]
                    price = int(p.find('p', attrs={'class': 'product-amount'}).contents[1].contents[0].split(' ')[1].replace(',',''))
                    amount = int(p.find('p', attrs={'class': 'product-amount'}).contents[3].contents[0][1:])
                    productsprice = productsprice + price
                    products.append([name,price,amount])

                shippingcost = orderprice - productsprice

                tmp.append([id,date,orderprice,productsprice,products])
            
            for i in reversed(tmp):
                orders.append(i)
            f.close()
        else:
            os.write(1, bytes('readfile error\n', 'utf-8'))

readDataHTML()
n = 0

def sheetWrite(row,col,data):
    x = 0
    for i in data:
        worksheet.write(col,row+x,i)
        x = x + 1

#worksheet.write(0, 0, 'Order')
#worksheet.write(0, 1, 'Product')
#worksheet.write(0, 2, 'Cost')
#worksheet.write(0, 3, 'Date')

sheetWrite(1,1,['Order','Product','Cost','Count','Date'])

for i in orders:
    n = n + 1
    #for x in range(0,len(i) - 1):
        #if x == len(i) - 2:
            #print(i[x],end='')
        #else:
            #print(i[x],end=' - ')
    #print('\n')
    #for x in range(0,len(i[len(i)-1])):
        #print('    ',i[len(i)-1][x])
    #print('\n')

n = 2
#[id,date,orderprice,productsprice,products]
for i in orders:
    sheetWrite(1, n, [i[0],'shipping',float((i[2]-i[3])/100),1,i[1]])
    n = n + 1
    for x in range(0,len(i[len(i)-1])):
        item = i[len(i)-1][x]
        #sheetWrite(1, n, [i[0],item[0],item[1],i[1]])
        sheetWrite(1, n, ['',item[0],float(item[1]/100),item[2],i[1]])
        n = n + 1
#sheetWrite(1,n + 1,['','','{=SUMPRODUCT(D3:D'+str(n)+';E3:E'+str(n)+')}'])

#worksheet.write_formula('D'+str(n+2),'=SUMPRODUCT(D3:D'+str(n)+';E3:E'+str(n)+')')

#print('\n\n\n',len(orders))
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 122)
worksheet.set_column('D:D', 7)
worksheet.set_column('E:E', 5)
worksheet.set_column('F:F', 20)
while True:
    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError as e:
        # For Python 3 use input() instead of raw_input().
        decision = input("Exception caught in workbook.close(): %s\n"
                             "Please close the file if it is open in Excel.\n"
                             "Try to write file again? [Y/n]: " % e)
        if decision != 'n':
            continue

    break