import zipfile, re
from xml.dom import minidom

class workbook(object):
    def __init__(self, filename, celltype = False):
        self.filename = filename
        self.ls = self.DOM(filename)
        self.Sheets = self.sheets(self.ls, celltype)
    
    
    
    def __repr__(self):
        return "<Workbook object [%s]@%s>" % (self.filename, self.Sheets)
    
    def __iter__(self):
        return iter(self.Sheets)
    
    class DOM():
        def __init__(self, filename):
            self.file = filename
        def __getitem__(self, file, r = False):
            if file[len(file)-1:len(file)] == "\%":
                file = file[0:len(file)-1]
                r = True
            filehandle = zipfile.ZipFile(self.file)
            if r:
                dom = filehandle
            else:
                dom = minidom.parseString(filehandle.read(file))
                filehandle.close
            return dom
        def __repr__(self):
            return self.file
    
    class sheets():
        sheets = {}
        i_sheets = []
        def __init__(self, ls, celltype):
            _sheets = ls["xl/workbook.xml"].documentElement.getElementsByTagName("sheets")[0]
            for sheet in _sheets.childNodes:
                
                obj = workbook.sheet(sheet._attrs['name'].value, ls, celltype)
                
                self.sheets[sheet._attrs['name'].value]= obj
                self.i_sheets.append(obj)
                self.__dict__[sheet._attrs['name'].value]= obj
        def keys(self):
            return self.sheets.keys()
        
        def __getitem__(self, key):
            if isinstance(key, int):
                if key >= len(self.i_sheets): key = len(self.i_sheets)-1
                return self.i_sheets[key]
            if key not in self.sheets.keys(): return None
            return self.sheets[key]
        
        def __iter__(self):
            return iter(self.i_sheets)
        
        def __len__(self):
            return len(self.sheets)
        
        def __repr__(self):
            return str(self.sheets)
        def __setattr__(self, attr, val):
            if attr not in self.__dict__.keys():
                if type(val) in [str]:
                    #Handles String Attributes
                    #This will be a new file
                    pass
                
            if isinstance(val, workbook.sheet):
                #Handles Workbook.sheet Attributes
                pass
                
    class sheet():
        sheet_dir = "xl/worksheets/"
        cells = {}
        def __init__(self, name, ls, celltype):
            self.name = name
            self.filename = self.sheet_dir + name.lower() + ".xml"
            self.dom = ls[self.filename]
            rows = self.dom.documentElement.getElementsByTagName("sheetData")[0].getElementsByTagName("row")
            if celltype:
                celltype = workbook.cell
            else:
                celltype = self.regcell
            for row in rows:
                cells = row.getElementsByTagName("c")
                for cell in cells:
                    name = cell.attributes.items()[0][1]
                    val = cell.getElementsByTagName("v")[0]._get_firstChild().nodeValue
                    cell = celltype(name, val)
                    self.cells[name] = cell
                    self.__dict__[name] = cell
        
        def __repr__(self):
            return  "<sheet '%s'>" % self.name
        
        def keys(self):
            return self.cells.keys()
        
        def __getattribute__(self, attr):
            if attr not in self.cells.keys():
                return None
            return self.__dict__[attr]
            
        def __getitem__(self, key):
            if key not in self.cells.keys():
                return None
            return self.cells[key]
        
        def __iter__(self):
            return iter(self.cells)
        
        def __len__(self):
            return len(self.cells)
        
        def regcell(self, name, val):
            if int(val):
                if float(val) - int(val) == 0: val = int(val)
                else: val = float(val)
            return val
    
    class cell(object):
        def __init__(self, name, val):
            self.name = name
            if int(val):
                if float(val) - int(val) == 0: val = int(val)
                else: val = float(val)
            self.val = val
        
        def __int__(self):
            return int(self.val)
        
        def __float__(self):
            return int(self.val)
        
        def __str__(self):
            return str(self.val)
        
        def __repr__(self):
            return "<Cell %s:'%s'>" % (self.name, self.val)
        
        def __add__(self, other):
            if isinstance(other, str) and isinstance(self.val, str):
                return self.val + other
            elif (isinstance(other, int) or isinstance(other, float)) and isinstance(self.val, float):
                return self.val + float(other)
            else:
                return str(self.val) + str(other)
        def __radd__(self, other):
            if isinstance(other, str) and isinstance(self.val, str):
                return other + self.val
            elif (isinstance(other, int) or isinstance(other, float)) and isinstance(self.val, float):
                return self.val + float(other)
            else:
                return  str(other)+str(self.val)
