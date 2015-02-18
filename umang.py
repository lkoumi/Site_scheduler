from flask import Flask, request, render_template, redirect, url_for, session, flash
import random
import csv
import xlsxwriter
from xlrd import open_workbook

app = Flask(__name__)
app.debug = True
app.config['SECRET_KEY'] = 'secret!'
app.secret_key = 'A0Zr98j/3yX R~XHH!jmN]LWX/,?RT'


# ------------------------------------------------------
@app.route('/')
def index1():
    return render_template('index.html')

@app.route('/duties')
def duties():
    return render_template('duties.html')
	
@app.route('/contact')
def contact():
	return render_template('contact.html')


@app.route('/duties1', methods=['POST'])
def duties1():
    getno()
    return redirect('/duties')

#------------------------------------------------------


ele = []
days = {}
main_d = {}
dut_no=[]


@app.route('/sendDuties', methods=['POST'])

def sendDuties():
    getval()
    date()
    print_file()
    return ''
    #return redirect('/duties')

def getno():
    t1 = request.form['t1']
    dut_no.append(t1)
    t2 = request.form['t2']
    dut_no.append(t2)
    t3 = request.form['t3']
    dut_no.append(t3)
    t4 = request.form['t4']
    dut_no.append(t4)
    t5 = request.form['t5']
    dut_no.append(t5)
    t6 = request.form['t6']
    dut_no.append(t6)
    t7 = request.form['t7']
    dut_no.append(t7)
    t8 = request.form['t8']
    dut_no.append(t8)
    t9 = request.form['t9']
    dut_no.append(t9)
    t10 = request.form['t10']
    dut_no.append(t10)
    t11 = request.form['t11']
    dut_no.append(t11)
    t12 = request.form['t12']
    dut_no.append(t12)
    t13 = request.form['t13']
    dut_no.append(t13)
    t14 = request.form['t14']
    dut_no.append(t14)


def getval():
    workbook=open_workbook('E:\code\Python\RC_sys\conn\List_faculty.xls')
    sheet1=workbook.sheet_by_index(0)
    data=[[sheet1.cell_value(r,1),sheet1.cell_value(r,2),sheet1.cell_value(r,3)]for r in range(1,sheet1.nrows)]
    #Integers from page or not?
    txt1 = request.form['txt1']
    txt2 = request.form['txt2']
    txt3 = request.form['txt3']
    txt4 = request.form['txt4']
    txt5 = request.form['txt5']
    txt6 = request.form['txt6']
    txt7 = request.form['txt7']
    txt8 = request.form['txt8']
    txt9 = request.form['txt9']
    txt10 = request.form['txt10']
    spaces = []
    for i in range(14):
        spaces.append('')
    k = 1
    for row in data:
        ele = [row[0], row[1], row[2]]
        if row[2] == "Senior Professor":
            ele.append(int(txt1))
            ele.append(int(txt1))
        elif row[2] == "Professor":
            ele.append(int(txt2))
            ele.append(int(txt2))
        elif row[2] == "Asso. Professor":
            ele.append(int(txt3))
            ele.append(int(txt3))
        elif row[2] == "Asst. Prof. Sel. Grade":
            ele.append(int(txt4))
            ele.append(int(txt4))
        elif row[2] == "Asst. Prof. Senior":
            ele.append(int(txt5))
            ele.append(int(txt5))
        elif row[2] == "Asst.Professor":
            ele.append(int(txt6))
            ele.append(int(txt6))
        elif row[2] == "Asst. Prof. Junior":
            ele.append(int(txt7))
            ele.append(int(txt7))
        elif row[2] == "Research Associate":
            ele.append(int(txt8))
            ele.append(int(txt8))
        elif row[2] == "Image Infotainment":
            ele.append(int(txt9))
            ele.append(int(txt9))
        elif row[2] == "Any other":
            ele.append(int(txt10))
            ele.append(int(txt10))
        ele.extend(spaces)
        main_d[int(k)] = ele
        k += 1


def assign():
    k = random.randint(1, len(main_d))
    #print main_d[k][4]
    if main_d[k][4] != 0:
        main_d[k][4] -= 1
        return k
    else:
        return assign()


def date():
    for i in range(14):
        day_l = []
        no_du = int(dut_no[i])
        #no_du=20
        #Input taken through page corresponding to date selected
        while no_du != 0:
            z = assign()
            day_l.append(z)
            if day_l.count(z) == 1:
                main_d[z][5 + i] = '1'
                no_du -= 1


def print_file():
    w_book = xlsxwriter.Workbook('E:\code\Python\RC_sys\conn\Disp.xlsx')
    w_sheet1=w_book.add_worksheet('First')
    r=2
    bold=w_book.add_format({'bold':1})
    m_form=w_book.add_format({'bold':1,'align':'center'})
    w_sheet1.merge_range('A1:D2','SITE Faculty',m_form)
    w_sheet1.merge_range('E1:F2','Duties',m_form)
    w_sheet1.write('A3','Sl.No.',bold)
    w_sheet1.write('B3','Gen.',bold)
    w_sheet1.set_column(1,1,8)
    w_sheet1.write('C3','Name',bold)
    w_sheet1.set_column(2,2,25)
    w_sheet1.write('D3','Designation',bold)
    w_sheet1.set_column(3,3,22)
    w_sheet1.write('E3','No. of Duties',bold)
    w_sheet1.write('F3','No. Alloted',bold)
    w_sheet1.set_column(4,5,12)
    w_sheet1.set_column(6,20,7)
    a=0
    for co in range(6,20):
    	if(co%2==0):
    		x=chr(65+a)
    		w_sheet1.merge_range(1,co,1,co+1,(x),m_form)
    		a+=1
    		w_sheet1.write(2,co,'I slot',bold)
    	else:
    		w_sheet1.write(2,co,'II slot',bold)
    for x in range(1,(len(main_d)+1)):
    	r+=1
    	main_d[x][4]=main_d[x][3]-main_d[x][4]
    	w_sheet1.write(r,0,x)
    	for i in range(1,(len(main_d[x])+1)):
    		item=main_d[x][i-1]
    		w_sheet1.write(r,i,item)
    w_book.close()
	#The above will give all the values of the cells of each row

if __name__ == '__main__':
    app.run()




