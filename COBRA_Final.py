# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'COBRA.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import pandas
from pandas import read_excel, DataFrame
import numpy as np 
import matplotlib.pyplot as plt
import datetime
import os.path 
matplotlib.use("Agg")

class Ui_Cobra(object):
    

    def setupUi(self, Cobra):
        Cobra.setObjectName("Cobra")
        Cobra.resize(250, 300)
        font = QtGui.QFont()
        font.setFamily("Century Schoolbook")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        Cobra.setFont(font)
        Cobra.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.centralwidget = QtWidgets.QWidget(Cobra)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.CIP_Label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Book Antiqua")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.CIP_Label.setFont(font)
        self.CIP_Label.setMouseTracking(True)
        self.CIP_Label.setObjectName("CIP_Label")
        self.verticalLayout.addWidget(self.CIP_Label)
        self.CIP_Report_Name = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Book Antiqua")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.CIP_Report_Name.setFont(font)
        self.CIP_Report_Name.setMouseTracking(True)
        self.CIP_Report_Name.setObjectName("CIP_Report_Name")
        self.verticalLayout.addWidget(self.CIP_Report_Name)
        self.PCS_Label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Book Antiqua")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.PCS_Label.setFont(font)
        self.PCS_Label.setMouseTracking(True)
        self.PCS_Label.setObjectName("PCS_Label")
        self.verticalLayout.addWidget(self.PCS_Label)
        self.PCS_Name = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Book Antiqua")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.PCS_Name.setFont(font)
        self.PCS_Name.setMouseTracking(True)
        self.PCS_Name.setObjectName("PCS_Name")
        self.verticalLayout.addWidget(self.PCS_Name)
        self.label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Book Antiqua")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.label.setFont(font)
        self.label.setMouseTracking(True)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.Sub_Task = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Book Antiqua")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.Sub_Task.setFont(font)
        self.Sub_Task.setMouseTracking(True)
        self.Sub_Task.setObjectName("Sub_Task")
        self.verticalLayout.addWidget(self.Sub_Task)
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Book Antiqua")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.pushButton.setFont(font)
        self.pushButton.setMouseTracking(True)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout.addWidget(self.pushButton)
        self.Data_Progress = QtWidgets.QProgressBar(self.centralwidget)
        self.Data_Progress.setProperty("value", 0)
        self.Data_Progress.setObjectName("Data_Progress")
        self.verticalLayout.addWidget(self.Data_Progress)
        Cobra.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(Cobra)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1120, 34))
        font = QtGui.QFont()
        font.setFamily("Book Antiqua")
        font.setBold(False)
        font.setWeight(50)
        self.menubar.setFont(font)
        self.menubar.setObjectName("menubar")
        self.menuMain_Window = QtWidgets.QMenu(self.menubar)
        self.menuMain_Window.setObjectName("menuMain_Window")
        self.menuBudget_Update = QtWidgets.QMenu(self.menubar)
        self.menuBudget_Update.setObjectName("menuBudget_Update")
        Cobra.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(Cobra)
        self.statusbar.setObjectName("statusbar")
        Cobra.setStatusBar(self.statusbar)
        self.menuBudget_Update.addSeparator()
        self.menubar.addAction(self.menuMain_Window.menuAction())
        self.menubar.addAction(self.menuBudget_Update.menuAction())

        self.retranslateUi(Cobra)
        QtCore.QMetaObject.connectSlotsByName(Cobra)

    def retranslateUi(self, Cobra):
        _translate = QtCore.QCoreApplication.translate
        Cobra.setWindowTitle(_translate("Cobra", "COBRA - Budgeting Tool"))
        self.CIP_Label.setText(_translate("Cobra", "CIP Report Document Title:"))
        self.PCS_Label.setText(_translate("Cobra", "PCS Document Title:"))
        self.label.setText(_translate("Cobra", "Search Subtasks for: "))
        self.pushButton.setText(_translate("Cobra", "Pull Data"))
        self.menuMain_Window.setTitle(_translate("Cobra", "Main Window"))
        self.menuBudget_Update.setTitle(_translate("Cobra", "Budget Update"))
        self.pushButton.clicked.connect(self.Assemble_PCS_compar)
        




#----------------------------------------------------------------PULL ALL DATA---------------------------------------------------------------------------------------

    def check_file_existence(self):
    	if os.path.isfile(self.cip_loc) == True and os.path.isfile(self.pcs_loc) == True:
    		return True
    	else:
    		return False

    def pull_PCS_data(self):
        self.Data_Progress.setValue(0)
        """CREATE PCS_compar DF: Creates data frame of [[Tab, task number n, money budgeted],....]"""
        tabs = ['Equip_Matl', 'Engr', 'Travel']
        tabs_totals = [11, 12, 8]
        df_tabs = [[], [], []]
        X2_val = []
        PCS_compar = []
        

        
        #creating dataframes from excel tabs
        for i in range(0, len(tabs)):
            df_tabs[i] = pandas.read_excel(self.pcs_loc, tabs[i], skiprows = 5)
            df_tabs[i] = df_tabs[i].values
        
        #Initial creation of PCS_compar dataframe - containing 3 elements: Tab, task number, and money budgeted 
        for j in range(0, 3):
            for i in range(0, len(df_tabs[j])):
                if type(df_tabs[j][i][5]) == str:
                    PCS_compar.append([tabs[j], df_tabs[j][i][5], df_tabs[j][i][tabs_totals[j]]])
                    #Consider changing the above code - took 1.3s without appending all at once... not as efficient (4.1s) using this method
                    X2_val.append(df_tabs[j][i][tabs_totals[j]])
        self.PCS_compar = PCS_compar
        
        return self.PCS_compar
        

    def CIP_report_DF(self):
        """Initial DF of CIP data"""
        cip_data = []
        cip_df = pandas.read_excel(self.cip_loc, skiprows = 13)
        cip_df = cip_df.values
        
        #removing NaN's from data
        for i in range(0, len(cip_df)):
            if type(cip_df[i][2]) == str:
                cip_data.append(cip_df[i])
        self.cip_df = cip_data
        
        return self.cip_df
        
    def pull_CIP_data(self):
        """DEPENDENCY = PCS_compar.....APPEND TO PCS_compar DF: Appends the CIP money spent per subtask into the PCS_compar datafram""" 
        PCS_compar = self.PCS_compar
        cip_df = self.cip_df
        spent = 0
        
        for j in range(0, len(PCS_compar)):
            spent = 0
            for i in range(0, len(cip_df)):
                if PCS_compar[j][1] in cip_df[i][2]:
                    spent = round(spent + cip_df[i][16] + cip_df[i][17], 2)
            
            if spent == 0:
                PCS_compar[j].append('Subtask not in this CIP report')
            else:
                PCS_compar[j].append(spent)
            
        self.PCS_compar = PCS_compar
        
        print(len(self.PCS_compar))
       
        return self.PCS_compar
        

    def add_ratio_spending(self):
        """DEPENDENCY = PCS_compar.....APPENDS to PCS_compar DF: Adds the ratio of spending relative to the CIP and PCS values"""
        PCS_compar = self.PCS_compar
        over_budget = 0
        progress = 0

        for line in range(0, len(PCS_compar)):
            progress = progress + len(PCS_compar)/100
            self.Data_Progress.setValue(progress)
            if type(PCS_compar[line][3]) != str:
                PCS_compar[line].append('{}%'.format(round(PCS_compar[line][3]/PCS_compar[line][2]*100, 2)))
                if PCS_compar[line][2] < PCS_compar[line][3]:
                    over_budget = over_budget + 1
            # else:
            #   PCS_compar[line].append('Not in CIP')
        self.PCS_compar = PCS_compar
        self.over_budget_tasks = over_budget

    def clean_pcs_compar(self):
        """CLEANS PCS_compar of any REPEAT LINES"""
        print(type(self.PCS_compar))
        PCS_compar = self.PCS_compar
        repeats = []
        repeats_check = []
        for i in range(0, len(PCS_compar)):
            for j in range(i + 1, len(PCS_compar)):
                if PCS_compar[i][1] == PCS_compar[j][1]:
                    PCS_compar[i][2] = PCS_compar[i][2] + PCS_compar[j][2]
                    if j not in repeats:
                        repeats.append(j)
                        repeats_check.append(PCS_compar[j])
        
        new_len = 0
        for line in repeats:
            

            del PCS_compar[line - new_len]
            new_len = new_len + 1
        
        
        self.PCS_compar = PCS_compar





        




#---------------------------------------------------------------MISC---------------------------------------------------------------------------------------


    def Total_X2(self):
        return sum(self.X2_val)

    def List_SubTasks(self):
        print(self.subtasks)


    def DF_Creation_TXT(self):
        """Generates a text document of PCS_compar DF"""
        column_s = ['Task Type', 'Task Number', 'Money Budgeted', 'Money Consumed', 'Spending Ratio']
        df_txt = pandas.DataFrame(data = self.PCS_compar, columns = column_s)
        self.df_txt = df_txt

        f = open('PCS_analysis.txt', 'w+')
        # f.write('{} of the subtasks are over-budget \n'.format(self.over_budget_tasks))
        f.write(str(df_txt))
        f.close()

#----------------------------------------------------------------SPECIAL FUNCTIONS FOR PULLING DATA YOU WANT---------------------------------------------------------------------------------------
    def sort_for_task(self, task_number):
        """Creates a list of lines from PCS_compar DF that match searched task number"""
        task_analysis = []
        for i in range(0, len(self.PCS_compar)):
            if task_number in self.PCS_compar[i][1]:
                task_analysis.append(self.PCS_compar[i])
        print(task_analysis)

        self.task_sorted = task_analysis

    def three_catagories_of_spending(self):
    	tasks = self.task_sorted
    	Equipment = 0
    	Engineering = 0
    	Travel = 0
    	print('HELLO')

    	for line in tasks:
    		if line[0] == 'Equip_Matl' and type(line[3]) != str:
    			print(line[3])
    			print(type(line[3]))
    			Equipment = Equipment + line[3]
    		if line[0] == 'Engr' and type(line[3]) != str:
    			Engineering = Engineering + line[3]
    		if line[0] == 'Travel' and type(line[3]) != str:
    			Travel = Travel + line[3]
    	print(Equipment, Engineering, Travel)





    def graph_my_tasks(self):
        """DEPENDENCY: task_sorted......Creates graph of tasks sorted for"""
        
        N = len(self.task_sorted)*2
        money_spent = []
        money_budgeted = []
        percent_consumed = []
        X_axis = []
        for i in range(0, len(self.task_sorted)):
            if type(self.task_sorted[i][3]) != str:
                money_spent.append((self.task_sorted[i][3]))
                percent_consumed.append(self.task_sorted[i][4])
                X_axis.append((self.task_sorted[i][1] + '({})'.format(percent_consumed[i])))
            else:
                money_spent.append((0))
                percent_consumed.append('')
                X_axis.append((self.task_sorted[i][1] + '(N/A)'))
            money_budgeted.append(self.task_sorted[i][2])
        ind = np.arange(0,N,2)  # the x locations for the groups
        width = 0.6      # the width of the bars


        if len(self.task_sorted) > 4:
            fig, ax = plt.subplots()
            rects1 = ax.barh(ind, money_spent, width, alpha = 0.75,color='b', edgecolor = 'black')
            rects2 = ax.barh(ind + width, money_budgeted, width, alpha = 0.7, color='orange', edgecolor = 'black')

            # add some text for labels, title and axes ticks
            ax.set_xlabel('USD - Currency')
            ax.set_title('Spending Report' + ' ' + str(datetime.date.today()))
            ax.set_yticks(ind + width/2)
            ax.set_yticklabels(X_axis)
            
            ax.legend((rects1[0], rects2[0]), ('Spent + Committed', 'Budgeted'))


            def autolabel(rects):
                counter = 0

                for rect in rects:
                    #print(rects)
                    width = int(rect.get_width())
                    if width < 100000:
                        x_loc = 1.01*width
                        horizontal_alignment = 'left'
                        color = 'black'
                    else:
                        x_loc = 0.98*width
                        horizontal_alignment = 'right'
                        color = 'black'
                    if width == 0:
                        width = ''
                    y_loc = rect.get_y() + rect.get_height()/3.
                    ax.text(x_loc,y_loc,('${}'.format(width)), ha=horizontal_alignment, va='center', color = color)
                    counter = counter + 1

            # columns = X_axis
            
            
            # columns = X_axis
            # row_labels = ['Spent', 'Budgetd', 'Ratio']
            # cell_text = [money_spent, money_budgeted, percent_consumed]
            
            # the_table = ax.table(cellText=cell_text,
   #                    rowLabels=row_labels,
   #                    colLabels=columns,
   #                    loc='bottom')


            

            autolabel(rects1)
            autolabel(rects2)
        

            plt.show()
        

        if len(self.task_sorted) < 4:
            

            fig, ax = plt.subplots()
            rects1 = ax.bar(ind + 1.2*width, money_budgeted, width, color='orange', edgecolor = 'black')
            rects2 = ax.bar(ind, money_spent, width, color='b', edgecolor = 'black')
            

            # add some text for labels, title and axes ticks
            ax.set_ylabel('USD - Currency')
            ax.set_title('Spending Report'+ ' ' +str(datetime.date.today()))
            ax.set_xticks(ind + 1.2*width/2)
            ax.set_xticklabels(X_axis)
            
            ax.legend((rects1[0], rects2[0]), ('Budgeted', 'Spent + Committed'))


            def autolabel(rects):
                counter = 0

                for rect in rects:
                    height = rect.get_height()
                    ax.text(rect.get_x() + rect.get_width()/2.,1.005*height,('${}'.format(round(height, 2))),ha='center', va='bottom')


            # columns = X_axis
            # row_labels = ['Spent', 'Budgetd', 'Ratio']
            # cell_text = [money_spent, money_budgeted, percent_consumed]
            
            # the_table = ax.table(cellText=cell_text,
   #                    rowLabels=row_labels,
   #                    colLabels=columns,
   #                    loc='bottom')




            autolabel(rects1)
            autolabel(rects2)
        
            self.Data_Progress.setValue(25)
            plt.show()


        
    def sort_tabs(self):
        equipment_matl = []
        for i in range(0, len(self.task_sorted)):
            if task_sorted[i][0] == 'Equip_Matl':
                equipment_matl.append([self.task_sorted])
                

            
    def Assemble_PCS_compar(self):
    	if self.CIP_Report_Name.text() != '' and self.PCS_Name.text() != '' and self.Sub_Task.text() !='':
    		self.cip_loc = self.CIP_Report_Name.text()
    		self.pcs_loc = self.PCS_Name.text()
    		task_number = self.Sub_Task.text()
    		if self.check_file_existence() == True:
		        self.pull_PCS_data()
		        self.CIP_report_DF()
		        self.pull_CIP_data()
		        self.add_ratio_spending()
		        self.clean_pcs_compar()
		        self.sort_for_task(task_number)
		        self.three_catagories_of_spending()
	        	self.graph_my_tasks()



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_Cobra()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())