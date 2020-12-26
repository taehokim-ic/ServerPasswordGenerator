# Server Password Generator

**_Generates or edits xlsx file containing passwords for servers that require regular password change_**

## Packages

```python
import random
from openpyxl import load_workbook
from openpyxl.styles import Font
```

## How to...

##### Create a PasswordList object, setting the row and column, generate PasswordList using passwordListGenerator()

```python
passwdlist = PasswordList()
passwdlist.setRows(10)
passwdlist.setColumns(3)
passwdlist.passwordListGenerator()
```

##### Loading file, setting font

```python
wb = load_workbook('password.xlsx')
ws = wb.active
ft = Font(name='Courier New', size=11)
```

##### Setting the values for each cell

```python
for i in range(passwdlist.rows):
    for j in range(passwdlist.columns):
        cell = chr(65+j) + str(i+1)   
        ws[cell].font = ft
        ws[cell] = passwdlist.passwordList[i][j]
```

> The following code assigns value to cells A1:C10; change if you wish.

## Details on Code

### 1. Generating Password

##### Within

```python
class Password():
```

##### Initializes password as empty string

```python
def __init__(self):
  
    self.password = ""
```

##### Checks whether certain character appears more than thrice

```python

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
```

##### Generates and returns password of given _length_, characters from ***'!'*** to ***'~'*** in ASCII Table

```python
def passwordGenerator(self, length):

    for i in range(length):
        passChr = chr(random.randint(33, 126))
        while (i == 0 and passChr == '='): # Solves bug when first character is '='
            passChr = chr(random.randint(33, 126))
        while not (self.lessThanThree(passChr)):
            passChr = chr(random.randint(33, 126))
        self.password += passChr

     return self.password

```

##### Sets password back to empty string

```python
def passwordClear(self):

    self.password = ""
```

### 2. Generating a set of _Passwords_

##### Within

```python
class PasswordList():
```

##### Creates instance

```python
def __init__(self):
      
    self.passwordList = []
    self.rows = 0      
    self.columns = 0
    self.password = Password()

```

##### Setter for _rows_ and _columns_

```python
def setRows(self, row):
    self.rows = row
  
def setColumns(self, column):
    self.columns = column
```

##### passwordListGenerator creates a list of password lists (row x column)

```python
def passwordListGenerator(self):
      
    passwords = []

    for i in range(self.rows * self.columns):
        self.password.passwordGenerator(9)
        while (self.password.password in passwords):
            self.password.passwordGenerator(9)
        passwords.append(self.password.password)
        self.password.passwordClear()
      
    for i in range(self.rows):
        _ = []
        for j in range(self.columns):
            _.append(passwords.pop())
        self.passwordList.append(_)

    return self.passwordList
```
