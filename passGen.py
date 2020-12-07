import random
from openpyxl import load_workbook

class Password():

    def __init__(self):
        
        self.password = ""

    def lessThanThree(self, passChr):

        try:
            index = self.password.index(passChr)

        except ValueError as e:
            return True

        try:
            index = self.password.index(passChr, index)
        
        except ValueError as e:
            return True

        return False

    def passwordGenerator(self, length):

        for i in range(length):
            passChr = chr(random.randint(33, 126))
            while not (self.lessThanThree(passChr)):
                passChr = chr(random.randint(33, 126))
            self.password += passChr

        return self.password

    def passwordClear(self):

        self.password = ""

class PasswordList():

    def __init__(self):
        
        self.passwords = []
        self.passwordList = []
        self.rows = 0
        self.columns = 0
        self.password = Password()

    def setRows(self, row):
        self.rows = row
    
    def setColumns(self, column):
        self.columns = column

    def passwordListGenerator(self):
        
        for i in range(self.rows * self.columns):
            self.password.passwordGenerator(9)
            while (self.password.password in self.passwords):
                self.password.passwordGenerator(9)
            self.passwords.append(self.password.password)
            self.password.passwordClear()
        
        for i in range(self.rows):
            _ = []
            for j in range(self.columns):
                _.append(self.passwords.pop())
            self.passwordList.append(_)

        return self.passwordList

passwdlist = PasswordList()
passwdlist.setRows(10)
passwdlist.setColumns(3)
print(passwdlist.passwordListGenerator())
