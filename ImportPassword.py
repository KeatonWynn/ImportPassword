import getpass
import pandas as pd


class ImportPassword:

    def __init__(self, File='Importfile'): # this enables you to pull in whatever your password file name is
        self.userHMF = getpass.getuser()
        self.myFile = "C:/Users/" + self.userHMF + "/ImportFiles/" + File + ".xlsx"
        #import_file = pd.ExcelFile(myFile, sheet=0)
        self.import_file = pd.ExcelFile(self.myFile)
        self.pwdf = self.import_file.parse(sheet=0)
        self.CSource = 0
        self.CUsername = 0
        self.CPassword = 0

    def getUser(self, source):
        try:
            row = self.pwdf.loc[self.pwdf['Source'] == source].index[0]
            value = self.pwdf.ix[row, 1]
            return str(value)
        except IndexError as e:
            print("Username and password not found. Please check spelling or check excel file in location below:")
            print(self.myFile)

    def getPassword(self, source):
        try:
            row = self.pwdf.loc[self.pwdf['Source'] == source].index[0]
            value = self.pwdf.ix[row, 2]
            if value != 'None':
                return str(value)
        except IndexError:
            print("Username and password not found. Please check spelling or check excel file in location below:")
            print(self.myFile)

