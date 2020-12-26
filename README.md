# Server Password Generator

**_Generates or edits xlsx file containing passwords for servers that need regular password change_**

## Packages

```python
import random
from openpyxl import load_workbook
from openpyxl.styles import Font
```


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
