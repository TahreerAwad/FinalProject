import pandas as pd
import numpy as np
from fpdf import FPDF
import matplotlib.pylab as plt
import matplotlib.patches as mpatches
import seaborn as sns
from PIL import Image
import os
from tqdm import tqdm # add bar progress

# get the path of file that can i read from
Input_path ="data.xlsx"

# Define My Main Object For inherit it .
class Main():
    # read all data in file
    df = pd.read_excel(Input_path)
    # read all data for students
    data = df.loc[df['Email'].notnull(), ['Email','Name', 'HW1','HW2', 'First', 'Second','final','total Grade' ]]
    #  get the specific data for each student
    all_names  = data['Name']
    all_Hw1 = data['HW1']
    all_Hw2 = data['HW2']
    all_First = data['First']
    all_Second = data['Second']
    all_final = data['final']
    all_Total = data['total Grade']
    # used for read a dynamic tasks on excel file
    hw1 = df.columns[3]
    hw2 = df.columns[4]
    First = df.columns[5]
    Second = df.columns[6]
    final = df.columns[7]
    # To check the value that is null or not
    def Check(self,value):
        if  value == 0 or str(value) == '':
            value="Miss or Did not submit"
            return value
        elif  value > 0:
            return str(int(value))


#  Make a new Class for Genrate a PDF
class PDF(FPDF):
    def __init__(self):
        super().__init__()
    # add footer that contain Page number
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', '', 12)
        self.cell(0, 8, f'Page {self.page_no()}', 0, 0, 'C')


# used for read a fake image that i save on it and remove it after that for all charts I have and for all Students
file_img= './images/image.png'
file_img2= './images/image2.png'
file_img3= './images/image3.png'

# Define My object
main = Main()
if not os.path.exists("output") or not os.path.exists("images") :
    os.mkdir("./output")
    os.mkdir("./images")
# looping on student in my data
for i in tqdm(range(len(main.data))):
    # check for folder that saved in !

    # # PDF creation for Each Student that inherit from PDF Class
    pdf = PDF() # Initializing object from my Class
    pdf.add_page() # Adding new page Page 1
    #  get the specific data for each student  in the first Class Called Main()
    name = main.all_names.iloc[i]
    Hw1 = main.all_Hw1.iloc[i]
    Hw2 = main.all_Hw2.iloc[i]
    First = main.all_First.iloc[i]
    Second = main.all_Second.iloc[i]
    final = main.all_final.iloc[i]
    Total = main.all_Total.iloc[i]

    # Title of Pdf Header for Each Student With his name
    pdf.set_fill_color(200,0,0)
    pdf.set_text_color(255,255,255)
    pdf.set_font("Arial", size=20, style="B")
    pdf.set_xy(5,5)
    text ='Details of Student : '+name  # get the name of each student
    pdf.cell(0, 12.5, txt=text, ln=1, align="C" , fill = True)

    # Text box of header
    pdf.set_fill_color(0,0,200)
    pdf.set_text_color(255,255,255)
    pdf.set_font("Arial", size=15, style='B')
    pdf.set_xy(5,20)
    pdf.cell(105, 10, txt='Grades in each of the course activities : ', ln=1, align="L",fill=True)


# __________________ First Task ___________#
    #• Student grades in each of the course activities (e.g., First, Second, HW1, Final, etc.)
    # Insert Values In table  box
    pdf.set_text_color(0,0,0)
    pdf.set_font("Arial", size=9, style="B")
    pdf.set_xy(5,35)
    pdf.cell(25+len(main.Check(Hw1)), 10, txt = "HW1 : " +main.Check(Hw1),  border=1, align="L")
    pdf.cell(25+len(main.Check(Hw2)), 10, txt = "HW2 : " +main.Check(Hw2),  border=1, align="L")
    pdf.cell(25+len(main.Check(First)), 10, txt = "First : " +main.Check(First),  border=1, align="L")
    pdf.cell(25+len(main.Check(Second)), 10, txt = "Second : " +main.Check(Second),  border=1, align="L")
    pdf.cell(25+len(main.Check(final)), 10, txt = "Final : " +main.Check(final),  border=1, align="L")


# __________________ Second Task ___________#
    # Header Of PieChart Digram
    pdf.set_fill_color(0,0,200)
    pdf.set_text_color(255,255,255)
    pdf.set_font("Arial", size=15, style ='B')
    pdf.set_xy(5,50)
    pdf.cell(130, 10, txt='Pie-chart showing the Weights of Course Activites:', ln=1, align="L",fill=True)

    # PieChart Digram
    my_data = [10, 10, 20,20,40]
    my_labels = 'Hw1', 'Hw2', 'First','Second','Final'
    plt.pie(my_data, labels=my_labels, autopct='%1.0f%%',wedgeprops={"edgecolor":"w",'linewidth': 2, 'antialiased': True})
    plt.title('Weights of Course Activities')
    plt.axis('equal')
    plt.savefig(file_img2, transparent=False,  facecolor='white', bbox_inches="tight")
    plt.close()
    pdf.image(file_img2, x = 0, y = 70, w = 200, h = 100, type = 'PNG')
    #pdf.add_page()

# __________________ Third Task ___________#
    # --------------------------------------------- #
    # A graphical representation (bar chart) of the student grades in the course activates
    # Header Of BarChart Digram
    pdf.set_fill_color(0,0,200)
    pdf.set_text_color(255,255,255)
    pdf.set_font("Arial", size=15, style='B')
    pdf.set_xy(5,170)
    pdf.cell(130, 10, txt='Bar-Chart For Student Grades in each Activates :', ln=1, align="L",fill=True)


    # Barchart Digram
    df = pd.DataFrame(dict(
    courses=['HW1', 'HW2', 'First', 'Second', 'all_final','Total Grade'],
    scored=[10, 10, 20, 20, 40,100],
    srudent_grades=[Hw1,Hw2, First, Second, final,Total]
    ))
    bar_plot1 = sns.barplot(x='courses', y='scored', data=df, color="red",width = .4 )
    bar_plot2 = sns.barplot(x='courses', y='srudent_grades', data=df, color="blue",width = .33)
    bar_plot2.set(ylabel='Grades',xlabel='')
    bar_plot2.set(title="Activites")
    red_patch = mpatches.Patch(color='red', label='Full Grade')
    blue_patch = mpatches.Patch(color='blue', label='Your Grade')
    plt.legend(handles=[red_patch, blue_patch])
    plt.tick_params(axis='x', rotation=90)
    plt.savefig(file_img, transparent=False,  facecolor='white', bbox_inches="tight")
    plt.close()
    pdf.image(file_img, x = 0, y = 190, w = 200, h = 100, type = 'PNG')
    pdf.add_page()


    # __________________ 4Th Task ___________#
    # • A chart showing the student his/her rank within the whole class
    # Text Header
    pdf.set_fill_color(0,0,200)
    pdf.set_text_color(255,255,255)
    pdf.set_font("Arial", size=15, style='B')
    pdf.set_xy(5,10)
    text = "It's " + name + ' Rank in the Whole Class For all Students :'
    pdf.cell(130, 10, txt=text, ln=1, align="L",fill=True)

    # Barchart Digram
    file_loc = "./data.xlsx"
    df = pd.read_excel(file_loc)
    # df = pd.read_excel(file_loc, index_col=None, na_values=['total Grade'], usecols="I,I:II")
    total_grade = df.loc[df['Email'].notnull(), ['total Grade' ]]
    total_name = df.loc[df['Name'].notnull(), ['Name' ]]
    lists = []
    names = []
    for i in range(len(total_grade.values)):
        lists.append(int(total_grade.values[i]))
    for i in range(len(total_name.values)):
        names.append(str(total_name.values[i]))
    values = lists
    idx = names
    clrs = ['red' if (x == Total) else 'blue' for x in values ]
    plt.ylabel('Grades')
    plt.title("Whole Class")
    plt.bar(idx, values, color=clrs, width=.9)
    plt.xticks('')
    plt.ylim(0, 100)
    plt.savefig(file_img3, transparent=False,  facecolor='white', bbox_inches="tight")
    pdf.image(file_img3, x = 0, y = 40, w = 200, h = 100, type = 'PNG')
    plt.close()

    # Footer Of My Page
    pdf.set_text_color(0,0,0)
    pdf.set_font("Arial", size=15, style='B')
    pdf.set_xy(5,160)
    text = "Thank You "+ name +" For Attend This Course "
    pdf.cell(130, 10, txt=text, ln=1, align="L")
    pdf.set_xy(5,210)
    pdf.set_font("Arial", size=20, style='B')
    text2 = "With My Best Wishes ,,,"
    pdf.cell(200, 10, txt=text2, ln=1, align="C")
    # End >>

    # # Exporting PDFs
    pdf.output("./output/"+name +".pdf")
    # Remove Pic For each plt
    os.remove(file_img)
    os.remove(file_img2)
    os.remove(file_img3)
