import operator 
from bs4 import BeautifulSoup
from urllib.request import urlopen
import re
from urllib.error import HTTPError
from urllib.error import URLError
import sqlite3
import time
import xlsxwriter
import time
from sqlite3 import Error
'''check the url browser'''

print("_______________________SEO Tool to Analyze Live Web Pages_____________________________________ ")
urls=["https://www.quora.com/","https://www.amazon.in/","https://www.awwwards.com/sites/stereo-3","https://twitter.com/login?lang=en","https://www.flipkart.com/"]#1 steps------contain group of url
adding_items=[]
print(".....connect to internet..........")
hello=1
for url in urls:
    try:
        file_handle=urlopen(url)
    except HTTPError as e:               #handle httperror
        print(e)
    except URLError as e:                #handle urlerror
        print("This server could not be found!!")
        print("not connected")
    else:
        print("connected succesfull")
        
        print("connecting to peer.........")
        file_handle=file_handle.read().decode('utf-8')
        soup=BeautifulSoup(file_handle,"html.parser")
        for a in soup(["script","style"]):
            a.extract()
        text=soup.get_text()
        line=(line.strip()for line in text.splitlines())                           
        a=[]   
        for word in line:
            word=word.split()
            a.extend(word)

        join_group=str(1*" ").join(a)#again string
        join_group=join_group.lower()
        join_group = re.sub("^\d+\s|\s\d+\s|\s\d+$|#|[0-9]|>>>|\.|\|", "", join_group)  #ignore
        adding_items=join_group.split()# adding total website(html) in form of list
        adding_items1=adding__items+adding_item1

    '''  read file igo which ignore words in html '''            #2nd steps
    with open("igo.txt",'r') as file:
        reader=file.read()                            #read file igo
        spilter=reader.split()
        spilter=[n for n in adding_items if n not in spilter]     #update list remove ignore word
        print("____processsing........")
        letter={}
        for i in spilter:
            if i not in letter:
                letter[i]=1
            else:
                letter[i]+=1
        sorted_list={}
        sorted_list=dict(sorted(letter.items(),key=operator.itemgetter(1), reverse=True)[:20])
        #sorted  upto descending order


                                                #connected to sqlite3'''
    if hello==1:                                                  #check database it will create then again use perivous datbase Select
        try:
            conn=sqlite3.connect('TEST.db')
        except Error as e:
            print(e)
        else:
            
            print("CONNTECTED database success")
            conn.execute('''drop table if exists hello''')
            conn.execute('''CREATE TABLE HELLO
            (
                NAME     TEXT    NOT NULL,
                REPEATWORDS   INT  NOT NULL,
                url TEXT 
                
            );'''
            )
            print("table create sucess")
            for k,v in sorted_list.items():
                conn.execute('INSERT INTO HELLO(NAME,REPEATWORDS,url) VALUES(?,?,?)',(k,v,url));
            conn.commit()

            ''' excel '''
            print("succesful create database")
            hello=hello+1
    else:
        try:
            conn=sqlite3.connect('TEST.db')
        except Error as e:
            print(e)
        else:
            print("CONNTECTED database success")
            for k,v in sorted_list.items():
                conn.execute('INSERT INTO HELLO(NAME,REPEATWORDS,url) VALUES(?,?,?)',(k,v,url));
            conn.commit()
c=conn.execute("SELECT NAME,REPEATWORDS,url FROM HELLO")
print("Process........finally wait")
time.sleep(1)
workbook = xlsxwriter.Workbook('C:\\output.xlsx')
worksheet = workbook.add_worksheet()
chart=workbook.add_chart({"type":"column","subtype":"stacked"})
chart1=workbook.add_chart({"type":"column","subtype":"stacked"})
chart2=workbook.add_chart({"type":"column","subtype":"stacked"})
chart3=workbook.add_chart({"type":"column","subtype":"stacked"})
chart4=workbook.add_chart({"type":"column","subtype":"stacked"})
conn=sqlite3.connect('TEST.db')

#c=conn.execute("SELECT NAME,REPEATWORDS,url FROM HELLO")
for i,row in enumerate(c):
    print("words:",row[0])                  #print database
    print("Count:",row[1])
    print("url:",row[2])

    worksheet.write(i,0,row[0])
    worksheet.write(i,1,row[1])
    worksheet.write(i,2,row[2])
'''  chart and input data in excel'''
chart.add_series({"name":"words","categories":"=Sheet1!$A$1:$A$20","values": "=Sheet1!$A$1:$A$20", 'column': {"color": 'blue'}})
chart.add_series({"name":"count","values": "=Sheet1!$B$1:$B$20", 'column': {"color": 'green'}})
#chart.add_series({"values": "=Sheet1!$C$1:$C$20", 'column': {"color": 'black'}})
chart.set_x_axis({'name': 'WORD','name_font': {'bold': True, 'italic': True}})
chart.set_y_axis({'name': 'COUNT','name_font': {'bold': True, 'italic': True}})
chart.set_title({'name': '=Sheet1!$C$1'})
chart.set_legend({'font': {'size': 9, 'bold': True}})
chart1.add_series({"name":"words","categories":"=Sheet1!$A$21:$A$40","values": "=Sheet1!$A$21:$A$40", 'column': {"color": 'blue'}})
chart1.add_series({"values": "=Sheet1!$B$21:$B$40", 'column': {"color": 'green'}})
#chart1.add_series({"values": "=Sheet1!$B$21:$B$41", 'column': {"color": 'black'}})
chart1.set_x_axis({'name': 'WORD','name_font': {'bold': True, 'italic': True}})
chart1.set_title({'name': '=Sheet1!$C$21'})
chart1.set_legend({'font': {'size': 9, 'bold': True}})
chart1.set_y_axis({'name': 'COUNT','name_font': {'bold': True, 'italic': True}})
chart2.add_series({"name":"words","values": "=Sheet1!$A$41:$A$60", 'column': {"color": 'blue'}})
chart2.add_series({"name":"count","values": "=Sheet1!$B$41:$B$60", 'column': {"color": 'green'}})
#chart2.add_series({"values": "=Sheet1!$B$42:$B$62", 'column': {"color": 'black'}})
chart2.set_x_axis({'name': 'WORD','name_font': {'bold': True, 'italic': True}})
chart2.set_title({'name': '=Sheet1!$C$41'})
chart2.set_legend({'font': {'size': 9, 'bold': True}})
chart2.set_y_axis({'name': 'COUNT','name_font': {'bold': True, 'italic': True}})
chart3.add_series({"name":"words","categories":"=Sheet1!$A$63:$A$83","values": "=Sheet1!$A$63:$A$83", 'column': {"color": 'blue'}})
chart3.add_series({"name":"count","values": "=Sheet1!$B$63:$B$83", 'column': {"color": 'green'}})
#chart3.add_series({"values": "=Sheet1!$B$63:$B$83", 'column': {"color": 'black'}})
chart3.set_x_axis({'name': 'WORD','name_font': {'bold': True, 'italic': True}})
chart3.set_title({'name': '=Sheet1!$C$61'})
chart3.set_legend({'font': {'size': 9, 'bold': True}})
chart3.set_y_axis({'name': 'COUNT','name_font': {'bold': True, 'italic': True}})
chart4.add_series({"name":"words","categories":"=Sheet1!$A$84:$A$104","values": "=Sheet1!$A$84:$A$104", 'column': {"color": 'blue'}})
chart4.add_series({"name":"count","values": "=Sheet1!$B$84:$B$104", 'column': {"color": 'green'}})
#chart4.add_series({"values": "=Sheet1!$B$84:$B$105", 'column': {"color": 'black'}})
chart4.set_x_axis({'name': 'WORD','name_font': {'bold': True, 'italic': True}})
chart4.set_title({'name': '=Sheet1!$C$81'})
chart4.set_legend({'font': {'size': 9, 'bold': True}})
chart4.set_y_axis({'name': 'COUNT','name_font': {'bold': True, 'italic': True}})
worksheet.insert_chart("K7",chart)
worksheet.insert_chart("S7",chart1)
worksheet.insert_chart("AA7",chart2)
worksheet.insert_chart("AI7",chart3)
worksheet.insert_chart("AR7",chart4)
conn.commit()
workbook.close()
print("okay complete")










