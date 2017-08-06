## This sample code demostrates how to extract data from an excel file and
## find the regression between two features selected by user
## Uses pandas and numpy for statistically operations
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import messagebox

class Regression:

    def __init(self):
        pass

    def findR(self,f_name,s_name):
            file = pd.ExcelFile(f_name)
            self.sheet_data=file.parse(s_name)
            print(self.sheet_data)
            headers = self.sheet_data[:]     # extract header names
            cnt=0;
            for i in headers:
                print(cnt,'',i)         # Print header names
                cnt=cnt+1
            x=input('Select X axis(enter no):: ')
            y=input('Select Y axis(enter no):: ')
            X_cnt=len(self.sheet_data.ix[:,int(x)])
            X_cnt = X_cnt - 2
            Y_cnt=len(self.sheet_data.ix[:,int(y)])
            Y_cnt = Y_cnt - 2
            xl=np.array(self.sheet_data.ix[1:X_cnt,int(x)])
            yl=np.array(self.sheet_data.ix[1:Y_cnt,int(y)])

            # converting in a numpy array
            N=X_cnt

            # Now creating the variables for Regression coeffient
            a=np.sum(np.multiply(xl,yl))
            b=np.sum(xl)
            c=np.sum(yl)
            d=np.sum(np.square(xl))
            e=np.sum(np.square(yl))

            # for plotting



            # R is regression coefficient
            R = float((N*a - b*c)/((N*d-b*b)*(N*e-c*c))**(1/2.0))

            # Call for fit line
            self.FitLine(xl,yl,b,y,N,R)

            return(R)

    def FitLine(self,xl,yl,b,y,N,R):
        mx=np.mean(xl)
        my=np.mean(yl)
        sx =np.std(xl) # std deviation of x
        sy =np.std(yl) # std deviation of y
        bo=R*(sy/sx)     # slope
        b1 = my - bo*mx  # intercept
        x=[]
        y=[]
        for i in range(0,10):
            x.append(i)
            y.append((b1 + bo*i))
        plt.subplot(2,1,1)
        plt.scatter(xl,yl,label='learning data')
        plt.plot(x,y,label='Fit line')

        plt.subplot(2,1,2)
        plt.plot(x,y,label='Fit line')


        plt.show()





def Display(regress):
    messagebox.showinfo('Regression Coefficient','The Regression Coefficient is '+str(regress))


if __name__ == "__main__":
    rg = Regression()
    f_name=input('Enter the Filename(with extension):: ')
    s_name=input('Enter the Sheet name(exactly same):: ')
    regress = rg.findR(f_name,s_name)
    Display(regress)
