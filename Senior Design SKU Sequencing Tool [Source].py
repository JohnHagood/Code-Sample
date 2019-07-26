#Senior Design SKU Scheduling
#If any questions arise, please contact johnhagood@gmail.com

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from csv import reader
from math import *
import xlsxwriter


#####################
##BEGIN READ RECIPE##
#####################

f=open("SKUMasterRecipe.csv")
reader=csv.reader(f)
next(reader)
next(reader)
SKURecipe={}
for row in reader:
    try:
        SKURecipe[int(row[0])]=[float(row[4]),float(row[5])]#length, thickness
    except:
        pass
f.close()

###################
##END READ RECIPE##
###################


try:##Read in regression weights, this file allows changes to model weights to be permanent
    f=open("Regression Weights.csv")
    reader=csv.reader(f)
    next(reader)
    for row in reader:
        thicknessWeight=float(row[0])
        lengthWeight=float(row[1])
        intercept=float(row[2])
    f.close()
except:
    messagebox.showerror("Error!","There was an issue with your 'Regression Weights.csv' file.\n Please ensure file is in correct format")

        
###########################
##BEGIN OPTIMIZATION CODE##
###########################
"""
    Our algorithm work by building an Eulerian tree by traversing the tree and when a node with degree > 1 is found,
    minimum spanning tree with unvisited nodes is created and eulerian tour is found for this tree
    the tour is then added to previously found tour
    """
parent={}

##Kruskal's Algorithm Implementation
def makeSet(vertex):
    parent[vertex]=vertex
def find(vertex):
    if parent[vertex]!=vertex:
         parent[vertex] =find(parent[vertex])
    return parent[vertex]
def union(vertexA,vertexB):
    xRoot = find(vertexA)
    yRoot = find(vertexB)
    parent[xRoot]=yRoot
def kruskal(graph):
    for vertex in graph["vertices"]:
        makeSet(vertex)
    minimum_spanning_tree=set()
    edges=list(graph["edges"])
    edges.sort()
    for edge in edges:
        if find(edge[1])!=find(edge[2]):
            union(edge[1],edge[2])
            minimum_spanning_tree.add(edge)
    return sorted(minimum_spanning_tree)
####End Kruskal's####


        
####Eulerian Tour Search####          
def findNext(vertices, current):##Find the next Node
    if vertices[current][1]>1:## if degree is higher than 1, return edge with lowest weight
        minEdge=vertices[current][2][0]
        
        for edge in vertices[current][2][1:]:
            if minEdge[0]>edge[0] and (vertices[edge[1]][0]==0 or vertices[edge[2]][0]==0):
                minEdge=edge
            if minEdge[1]!=current:
                minNode=minEdge[1]
            else:
                minNode=minEdge[2]
                
        return (minNode,True)
    else:## if degree is equal to 1, return edge 
        if vertices[current][2][0][1]!=current:
            nextNode= vertices[current][2][0][1]
        else:
            nextNode= vertices[current][2][0][2]
        return (nextNode,True)

def describeTree(graph,tree):##Describe all vertices in tree as [visited {1 if visited, 0 otherwise}, Degree, List of Edges]
    vertices={}
    for vertex in graph["vertices"]:
        vertices[vertex]=[0,0,[]]
    for edge in tree:
        vertices[edge[1]][1]+=1
        vertices[edge[2]][1]+=1
        if edge[1] in vertices and edge[2] in vertices:
            vertices[edge[1]][2].append(edge)
            vertices[edge[2]][2].append(edge)
    return vertices
        
def eulerianTour(graph,tree, start):
    """
    Builds Eulerian tree by traversing the tree and when a node with degree > 1 is found,
    minimum spanning tree with unvisited nodes is created and eulerian tour is found for this tree
    the tour is then added to previously found tour
    """
    vertices=describeTree(graph,tree)
    tour=[]
    tour.append(start)
    current=start
    while len(tour)<len(list(vertices)):
        vertices[current][0]=1
        nextNode=findNext(vertices,current)
        current=nextNode[0]
        if nextNode[1]:
            newGraph={"vertices":[],"edges":set([])}
            for skuA in vertices:
                if vertices[skuA][0]==0:
                    newGraph["vertices"].append(skuA)
                for skuB in vertices:
                    if skuA != skuB and vertices[skuB][0]==0 and vertices[skuA][0]==0:
                        recipeA=SKURecipe[skuA]
                        recipeB=SKURecipe[skuB]
                        dLength=abs(recipeA[0]-recipeB[0])
                        dThick=abs(recipeA[1]-recipeB[1])
                        weight=thicknessWeight*dThick+lengthWeight*dLength
                        newGraph["edges"].add((weight,skuA,skuB))
            newTree=kruskal(newGraph)
            newTour=eulerianTour(newGraph,newTree, current)
            tour=tour+newTour
        else:
            tour.append(current)
        
    return tour
####End Eulerian Tour search####
            
        
                

def findSequence(orders):
    ##Run Approximation algorithm code
    
    ##create graph where each edge is weighted by the difference in thickness and length between each subsequent SKU
    graph={"vertices":[],"edges":set([])}
    skuList=[]
    for skuA in orders:
        graph["vertices"].append(skuA)
        skuList.append([skuA,orders[skuA], SKURecipe[skuA][0],SKURecipe[skuA][1]])
        for skuB in orders:
            if skuA != skuB:
                recipeA=SKURecipe[skuA]
                recipeB=SKURecipe[skuB]
                dLength=abs(recipeA[0]-recipeB[0])
                dThick=abs(recipeA[1]-recipeB[1])
                weight=1.822*dThick+.00384*dLength##Weight determined by regression model
                graph["edges"].add((weight,skuA,skuB))
                
        
        
    skuList.sort(key=lambda x: (x[3],x[2]))##start with lowest thickness, length
    finalSequence=[skuList[0][0]]
    currentSku=skuList[0][0]
    tree=kruskal(graph)
    sequence=eulerianTour(graph,tree, currentSku)
    
    return sequence
#########################
##END OPTIMIZATION CODE##
#########################







#############
##BEGIN GUI##
#############

        
class SKUSchedule:
    def __init__(self, win):##Create title screen
        self.photo=PhotoImage(file='Company Logo.gif')
        self.pho=Label(win, image=self.photo, bg="white")
        self.pho.grid(row=0, column=0, columnspan=4, padx=30, pady=30)
        self.b1=Button(win, text="Import", command=self.importing, bg="white")
        self.b1.grid(row=1, column=0, sticky=E+W, padx=30, pady=20)
        self.runModelButton=Button(win, text="Run Model", command=self.runModel, bg="white",state=DISABLED)
        self.runModelButton.grid(row=1, column=1, sticky=E+W, padx=30, pady=20)
        self.exportModelButton=Button(win, text="Export to Excel", command=self.exporting, bg="white",state=DISABLED)
        self.exportModelButton.grid(row=1, column=2, sticky=E+W, padx=30, pady=20)
        self.changeRegressionButton=Button(win, text="Change Model Weights", command=self.changeWeights, bg="white")
        self.changeRegressionButton.grid(row=1, column=3, sticky=EW, padx=30, pady=20)
        self.b4=Button(win, text="Exit", command=self.exiting, bg="white")
        self.b4.grid(row=1, column=4, sticky=E+W, padx=30, pady=20)
        self.l=Label(win, text="Imported CSV File:", bg="white")
        self.l.grid(row=2, column=0)
        self.file1=StringVar()
        self.file1.set("")
        self.e=Entry(win, textvariable=self.file1, state='readonly', width='90', bg="white")
        self.e.grid(row=2, column=1, columnspan=3, pady=15, padx=15)


####################
##HELPER FUNCTIONS##
####################

    
    def calculateObjValue(self,sequence): ##calculate regression formula for inputted sequence
        thicks=[]
        lengths=[]
        for sku in sequence:
            recipe=SKURecipe[sku]
            thicks.append(recipe[1])
            lengths.append(recipe[0])
        dThick=0
        dLength=0
        for i in range(1,len(lengths)):
            Tchange=abs(thicks[i]-thicks[i-1])
            Lchange=abs(lengths[i]-lengths[i-1])
            dThick+=Tchange
            dLength+=Lchange
        objValue=thicknessWeight*dThick+lengthWeight*dLength+intercept##Regression Equation
        boards=(4539.01*exp((objValue-0.965219)/0.641844)+219.556)/(1+exp((objValue-0.965219)/0.641844))##Johnson Transformation Reversal
        return dThick,dLength,boards

    def parseQuantity(self, quantity): ##Parses quantity input for commas
        try:
            try:
                return float(quantity)
            except:
                if "," in quantity:
                    quantity=quantity.replace(",","")
                    return float(quantity)
                else:
                    messagebox.showerror("Error!","""Sorry, it seems you have input an incorrent number type into the quantity column.\n
                                                    Please make sure all numbers are in the format of ####### or ###,###""")
        except:
            messagebox.showerror("Error!","""Sorry, it seems you have input an incorrent number type into the quantity column.\n
                                            Please make sure all numbers are in the format of ####### or ###,###""")
    def exporting(self): ##export to excel format
        fileList=[]
        for skus in self.finalSequence:
            for order in self.productOrders[skus]:
                fileList.append([order[0],skus,order[1]])
        file=filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel Files","*.xlsx"),("All Files","*.*")))
        if ".xlsx" not in file:
            messagebox.showerror("Error","Export files must be in xlsx format")
            return
        workbook = xlsxwriter.Workbook(file)
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 20)
        worksheet.set_column('C:C', 20)
        worksheet.write(0,0,"Product ID")
        worksheet.write(0,1,"Sku #")
        worksheet.write(0,2,"Quantity of Boards")
        for i,file in enumerate(fileList):
            worksheet.write(i+1,0,file[0])
            worksheet.write(i+1,1,file[1])
            worksheet.write(i+1,2,file[2])

        workbook.close()
        messagebox.showinfo("Success!","You have successfully exported the optimal run.")

    def exportingCustom(self): ## export custom run sequence into excel format
        fileList=[]
        for skus in self.finalOrderedList:
            for order in self.productOrders[skus]:
                fileList.append([order[0],skus,order[1]])
        file=filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel Files","*.xlsx"),("All Files","*.*")))
        if ".xlsx" not in file:
            messagebox.showerror("Error","Export files must be in xlsx format")
            return
        workbook = xlsxwriter.Workbook(file)
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 20)
        worksheet.set_column('C:C', 20)
        worksheet.write(0,0,"Product ID")
        worksheet.write(0,1,"Sku #")
        worksheet.write(0,2,"Quantity of Boards")
        for i,file in enumerate(fileList):
            worksheet.write(i+1,0,file[0])
            worksheet.write(i+1,1,file[1])
            worksheet.write(i+1,2,file[2])

        workbook.close()
        messagebox.showinfo("Success!","You have successfully exported your custom run.")

    def importing(self): ## import order file
        try:
            file=filedialog.askopenfilename()
            self.file1.set(file)
            self.fileName1= file
            f= open(file, "r")
            reader= csv.reader(f)
            data=[]
            for row in reader:
                if row[1] != "" and row[2] != "":
                    data=data+[row]
            
            data=data[1:len(data)]
            self.datalist = data #self.datlist orders in list of lists
            
            for row in self.datalist:
                
                row[1]=int(row[1])
                row[2]=self.parseQuantity(row[2])
                
                if row[2] == None:
                    return
                                           
            self.orderInput = {} # self.orderInput is combined order quanity for skus {sku: sum orders, sku: sum orders}
            self.productOrders = {}
            for item in self.datalist:
                if item[1] in self.orderInput:
                    self.orderInput[item[1]]=float(self.orderInput[item[1]])+float(item[2])
                    self.productOrders[item[1]].append((item[0],item[2]))
                    
                else:
                    self.orderInput[item[1]]=float(item[2])
                    self.productOrders[item[1]]=[(item[0],item[2])]
            self.numskus = len(self.orderInput)
            self.runModelButton.config(state=NORMAL)
        except:
            messagebox.showerror("Error!","Oops! there seems to be a mistake in your file.\n Please ensure it is in the correct format and try again.")
    def estimateDefects(self, index, sequence):
        prev=SKURecipe[sequence[index-1]]
        curr=SKURecipe[sequence[index]]
        lengthDifference=abs(curr[0]-prev[0])
        thickDifference=abs(curr[1]-prev[1])
        y=thicknessWeight*thickDifference+lengthWeight*lengthDifference##Regression Equation Weights
        return y
    def makeWeightChanges(self):
        thicknessWeight=self.newThick.get()
        lengthWeight=self.newLength.get()
        intercept=self.newIntercept.get()
        f=open("Regression Weights.csv","w", newline="")
        writer=csv.writer(f)
        writer.writerow(["Thickness Weight","Length Weight", "Intercept"])
        writer.writerow([thicknessWeight, lengthWeight, intercept])
        f.close()
        messagebox.showinfo("Success!","You have successfully updated the model weights, please restart your software before re-running your model.")
        self.changeWeightsWin.withdraw()
                        

########################
##END HELPER FUNCTIONS##
########################
                

########################
##GUI SCREEN FUNCTIONS##
########################


    def runModel(self):## create optimal run frame

        if len(self.orderInput)==0:
            return
        self.f=Frame(win, bg='white')
        self.f.grid(row=3, column=0, columnspan=4)
        self.l2=Label(self.f, text="Optimal Solution", bg='white')
        self.l2.grid(row=0, column=0, columnspan=2, sticky=E+W)
        self.title1=Label(self.f, text="SKU", bg='white')
        self.title1.grid(row=1, column=0, sticky=E+W)
        self.title2=Label(self.f, text="Quantity", bg='white')
        self.title2.grid(row=1, column=1, sticky=E+W)
        self.title3=Label(self.f, text="Production Run Order", bg='white')
        self.title3.grid(row=1, column=2, sticky=E+W)
        self.title4=Label(self.f, text="Est. Defects This Run", bg='white')
        self.title4.grid(row=1, column=3, sticky=EW)
        orders=self.orderInput
        sequence=findSequence(orders)
        self.finalSequence=sequence
        self.indexList = []
        fractionsList=[]
        for i in range(len(sequence)): #skus and quantity output
            self.s= Label(self.f, text=sequence[i],bg='white')
            self.s.grid(row=i+2, column=0, sticky=E+W)
            self.o= Label(self.f, text=self.orderInput[sequence[i]],bg='white')
            self.o.grid(row=i+2, column=1, sticky=E+W)

            self.r=Label(self.f, text=i+1,bg='white')
            self.r.grid(row=i+2, column=2, sticky=E+W)
            self.indexList.append(i+1)
            if i>0:
                fractionsList.append(self.estimateDefects(i, sequence))
        
        boards=self.calculateObjValue(sequence)
        total=sum(fractionsList)
        for i in range(1,len(sequence)):
            self.defects=Label(self.f, text="{:.1f}".format((fractionsList[i-1]/total)*boards[2]), bg="white")
            self.defects.grid(row=i+2, column=3, sticky=E+W)


        self.thickshift=Label(self.f,bg='white', text="Net Thickness Shift (in.): {:.3f}".format(boards[0]))
        self.thickshift.grid(row=2, column=4, sticky=E+W)
        self.lenshift=Label(self.f,bg='white', text="Net Length Shift (in.): {:.1f}".format(boards[1]))
        self.lenshift.grid(row=3, column=4, sticky=E+W)
        self.estdef=Label(self.f,bg='white', text="Est. Defective Boards: {:.1f}".format(boards[2]))
        self.estdef.grid(row=5, column=4, sticky = E+W)
    

        self.editButton = Button(self.f,bg='white', text="Edit Optimal Model", command=self.editModel)
        self.editButton.grid(row=6, column=4, sticky= E+W)
        self.exportModelButton.config(state=NORMAL)

    def editModel(self):## window with edit entries
        win.deiconify()
        self.editModelPage = Toplevel()

        self.f=Frame(self.editModelPage, bg='white')
        self.f.grid(row=3, column=0, columnspan=4)
        self.l2=Label(self.f, text="Optimal Solution", bg='white')
        self.l2.grid(row=0, column=0, columnspan=2, sticky=E+W)
        self.title1=Label(self.f, text="Previous Run Order", bg='white')
        self.title1.grid(row=1, column=0, sticky=E+W)
        self.title1=Label(self.f, text="SKU", bg='white')
        self.title1.grid(row=1, column=1, sticky=E+W)
        self.title2=Label(self.f, text="Quantity", bg='white')
        self.title2.grid(row=1, column=2, sticky=E+W)
        self.title3=Label(self.f, text="Test Run Order", bg='white')
        self.title3.grid(row=1, column=3, sticky=E+W)
        sequence=self.finalSequence
        self.skuEditDictionary={}
        for i in range(len(sequence)): 
            self.previousOrderLabel=Label(self.f,text=i+1, bg='white')
            self.previousOrderLabel.grid(row=i+2, column=0, sticky=E+W)
            self.skuEditDictionary[sequence[i]]=StringVar(self.f)
            self.s= Label(self.f, text=sequence[i], bg='white')
            self.s.grid(row=i+2, column=1, sticky=E+W)
            self.o= Label(self.f, text=self.orderInput[sequence[i]], bg='white')
            self.o.grid(row=i+2, column=2, sticky=E+W)

            self.r=Entry(self.f, textvariable=self.skuEditDictionary[sequence[i]])
            self.r.grid(row=i+2, column=3, sticky=E+W)
           

        backButton = Button(self.editModelPage, text="Back", command=self.backToResults, bg='white')
        backButton.grid(row=7,column=0, columnspan=4, sticky = E+W)

        makeChangesButton = Button(self.editModelPage, text="Confirm Changes", command=self.makeChanges, bg='white')
        makeChangesButton.grid(row=8,column=0, columnspan=4, sticky = E+W)

    def makeChanges(self): ##window with changes made
        
        try:
            ## Take in edited sku sequence and display on new window
            newOrderedList = []
            for keys,values in self.skuEditDictionary.items():
                entry=int(values.get())
                newOrderedList.append([keys, int(values.get())])
                if entry not in range(len(self.skuEditDictionary)+1):
                    messagebox.showerror("Error!","All numbers must be in the range 1-{}".format(len(self.skuEditDictionary)))
                    return
                

            newOrderedList.sort(key=lambda x: x[1])



            finalOrderedList = []
            runOrderList=[]
            for subItem in newOrderedList:
                finalOrderedList.append(subItem[0])
                runOrderList.append(subItem[1])
            self.finalOrderedList=finalOrderedList
            runOrderDict={}
            for i in range(len(self.finalOrderedList)+1):
                runOrderDict[i]=0
            for item in runOrderList:
                runOrderDict[item]+=1
                if runOrderDict[item]>1:
                    messagebox.showerror("Error!","No run order may be chosen more than once.")
                    return

        except:
            messagebox.showerror("Error!","There has been an error with your entries, check the format and try again.")
            return
        self.editModelPage.withdraw()
        self.changedPage = Toplevel()
        
        self.f=Frame(self.changedPage, bg='white')
        self.f.grid(row=3, column=0, columnspan=4)
        self.l2=Label(self.f, text="Custom Sequence", bg='white')
        self.l2.grid(row=0, column=0, columnspan=2, sticky=E+W)
        self.title1=Label(self.f, text="SKU", bg='white')
        self.title1.grid(row=1, column=0, sticky=E+W)
        self.title2=Label(self.f, text="Quantity", bg='white')
        self.title2.grid(row=1, column=1, sticky=E+W)
        self.title3=Label(self.f, text="Production Run Order", bg='white')
        self.title3.grid(row=1, column=2, sticky=E+W)

        
        fractionsList=[]
        for i,sku in enumerate(finalOrderedList):
            self.s= Label(self.f, text=sku, bg='white')
            self.s.grid(row=i+3, column=0, sticky=E+W)
            self.o= Label(self.f, text=self.orderInput[sku], bg='white')
            self.o.grid(row=i+3, column=1, sticky=E+W)

            self.r=Label(self.f, text=i+1, bg='white')
            self.r.grid(row=i+3, column=2, sticky=E+W)
            if i>0:
                fractionsList.append(self.estimateDefects(i, finalOrderedList))
        
        boards=self.calculateObjValue(finalOrderedList)
        total=sum(fractionsList)
        for i in range(1,len(finalOrderedList)):
            self.defects=Label(self.f, text="{:.1f}".format((fractionsList[i-1]/total)*boards[2]), bg="white")
            self.defects.grid(row=i+3, column=3, sticky=E+W)                       
        
        
        self.thickshift=Label(self.f, text="Net Thickness Shift (in.): {:.3f}".format(boards[0]), bg="white")
        self.thickshift.grid(row=2, column=4, sticky=E+W)
        self.lenshift=Label(self.f, text="Net Length Shift (in.): {:.1f}".format(boards[1]), bg="white")
        self.lenshift.grid(row=3, column=4, sticky=E+W)

        
        self.estdef=Label(self.f, text="Est. Defective Boards: {:.1f}".format(boards[2]), bg="white")
        self.estdef.grid(row=5, column=4, sticky = E+W)


        backButton = Button(self.f, text="Back", command=self.backFromChanged,bg="white")
        backButton.grid(row=7,column=4, columnspan=1, sticky = E+W)

        exportButton = Button(self.f, text="Export This Sequence", command=self.exportingCustom,bg="white")
        exportButton.grid(row=8,column=4, columnspan=1, sticky = E+W)

    def changeWeights(self):
        self.changeWeightsWin=Toplevel()
        self.thickLabel=Label(self.changeWeightsWin, text="New Thickness Weight:", bg="white")
        self.thickLabel.grid(row=0, column=0,sticky=EW)
        self.lengthLabel=Label(self.changeWeightsWin, text="New Length Weight:", bg="white")
        self.lengthLabel.grid(row=1, column=0, sticky=EW)
        self.interceptLabel=Label(self.changeWeightsWin, text="New Intercept Value:", bg="white")
        self.interceptLabel.grid(row=2, column=0, sticky=EW)
        self.newThick=StringVar()
        self.newLength=StringVar()
        self.newIntercept=StringVar()
        self.thickEntry=Entry(self.changeWeightsWin, textvariable=self.newThick, bg="white")
        self.thickEntry.grid(row=0, column=1, sticky=EW)
        self.lengthEntry=Entry(self.changeWeightsWin, textvariable=self.newLength, bg="white")
        self.lengthEntry.grid(row=1, column=1, sticky=EW)
        self.interceptEntry=Entry(self.changeWeightsWin, textvariable=self.newIntercept, bg="white")
        self.interceptEntry.grid(row=2, column=1, sticky=EW)
        self.confirmButton=Button(self.changeWeightsWin, text="Confirm Changes", command=self.makeWeightChanges)
        self.confirmButton.grid(row=3, column=0, columnspan=2, sticky=EW)
        
        
        





        

                                
    def exiting(self):
        win.withdraw()

    def backToResults(self):
        self.editModelPage.withdraw()
    def backFromChanged(self):
        self.changedPage.withdraw()
        self.editModelPage.deiconify()

                                  
        
    
win=Tk()
win.title('Company Production Scheduling Tool')
win.config(bg='white')
app= SKUSchedule(win)
win.mainloop()
