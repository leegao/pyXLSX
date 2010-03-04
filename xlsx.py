import zipfile, re
from xml.dom import minidom
from range_alpha import range_alpha
from sys import exc_info


class workbook(object):
    def __init__(self, filename, celltype = False):
        self.filename = filename
        self.ls = self.DOM(filename)
        self.Sheets = self.sheets(self.ls, celltype, self)
        
    def __repr__(self):
        return "<Workbook object [%s]@%s>" % (self.filename, self.Sheets)
    
    def __iter__(self):
        return iter(self.Sheets)
    
    def extend(self, fn):
        if not ('func_code' in dir(fn)): raise RuntimeError('Cannot Extend Workbook Formulas with Nonfunctions')
        key = fn.func_name
        self.__dict__[key.upper()]=fn
    
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
                _str = filehandle.read(file)
                dom = minidom.parseString(_str)
                filehandle.close
            return dom
        def __repr__(self):
            return self.file
    
    class sheets():
        sheets = {}
        i_sheets = []
        def __init__(self, ls, celltype, workbook):
            _sheets = ls["xl/workbook.xml"].documentElement.getElementsByTagName("sheets")[0]
            _shared = ls["xl/sharedStrings.xml"].documentElement.getElementsByTagName("si")
            for sheet in _sheets.childNodes:
                obj = workbook.sheet(sheet._attrs['name'].value, sheet._attrs['r:id'].value.replace("rId",""), ls, celltype, _shared, workbook)
                
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
        def __init__(self, name, id, ls, celltype, shared, workbook):
            self.name = name
            self.filename = self.sheet_dir + "sheet" + id + ".xml"
            self.dom = ls[self.filename]
            self.workbook = workbook
            rows = self.dom.documentElement.getElementsByTagName("sheetData")[0].getElementsByTagName("row")
            if not celltype:
                celltype = workbook.cell
            else:
                #Revert to legacy support
                celltype = self.regcell
            for row in rows:
                cells = row.getElementsByTagName("c")
                for cell in cells:
                    name = cell._attrs['r'].value
                    
                    _share = False
                    _fn = None
                    
                    if 't' in cell._attrs: _share = True
                    if cell.getElementsByTagName("f"): _fn = cell.getElementsByTagName("f")[0]._get_firstChild().nodeValue
                    
                    try:
                        val = cell.getElementsByTagName("v")[0]._get_firstChild().nodeValue
                    except:
                        #Blank cell
                        val = ""
                        
                    if _share:
                        val = shared[int(val)].getElementsByTagName("t")[0]._get_firstChild().nodeValue
                    
                    cell = celltype(name, val, self, _fn)
                    self.cells[name] = cell
                    self.__dict__[name] = cell
        
        def __repr__(self):
            return  "<Sheet '%s'>" % self.name
        
        def keys(self):
            return self.cells.keys()
        
        def __getattribute__(self, attr):
            if attr not in self.cells.keys():
                return None
            _return = self.__dict__[attr]
            #if isinstance(_return, workbook.cell): _return = _return.val
            return _return
            
        def __getitem__(self, key):
            if key not in self.cells.keys():
                return None
            _return = self.__dict__[key]
            #if isinstance(_return, workbook.cell): _return = _return.val
            return self.cells[key]
        
        def __iter__(self):
            return iter(self.cells)
        
        def __len__(self):
            return len(self.cells)
        
        def PARSE(self, match):
            return "workbook."+match.group().upper()
        
        def REPLACE(self, match):
            tokens = match.group().split(":")
            #Structure:
            #    A+1+
            #Cases:
            #    A1:A2 - Vertical
            #    A1:B1 - Horizontal
            #    A1:B2 - Rectangular
            #    -> then V
            alpha = re.compile(r"[a-z|A-Z]+")
            numer = re.compile(r"[0-9]+")
            
            A0 = alpha.search(tokens[0]).group().upper()
            A1 = alpha.search(tokens[1]).group().upper()
            N0 = int(numer.search(tokens[0]).group())
            N1 = int(numer.search(tokens[1]).group())
            
            #Case 1: Same alpha
            if A0 == A1:
                #Range in numer
                _range = [A0+str(n) for n in range(N0, N1+1)]
                _ret = [self[o].val for o in _range]
                return str(_ret)
            
            #Case 2: Same numer
            if N0 == N1:
                #Range in alpha
                _range = [a+str(N0) for a in range_alpha(A0, A1)]
                _ret = [self[o].val for o in _range]
                return str(_ret)
            
            #Case 3: Else
            
            ###Unimplemented Yet###
            
            return tokens
        def interpolate(self, text):
            rn = re.compile(r"[a-z|A-Z]+[0-9]+\:*[a-z|A-Z]+[0-9]+\:*")
            fn = re.compile(r"(?<=\(|\))?[a-z|A-Z]*(?=\()")
            return rn.sub(self.REPLACE, fn.sub(self.PARSE, text))
        
        def regcell(self, name, val, sheet, fn=None):
            #Legacy
            try:
                if float(val):
                    val = float(val)
                    if val - int(val) == 0: val = int(val)
            except ValueError:
                pass
            return val
    
    class cell(object):
        def __init__(self, name, val, sheet, fn=None):
            self.name = name
            try:
                if float(val):
                    val = float(val)
                    if val - int(val) == 0: val = int(val)
            except ValueError:
                pass
            
            self.fn = fn
            self.sheet = sheet
            self.val = val
        def parse(self):
            if self.fn:
                return self.sheet.interpolate(self.fn)
        def evaluate(self, strict=True):
            result = self.parse()
            if not result:
                raise RuntimeError('Cannot Evaluate Unsyntaxical Expressions')
            try:
                self.val = eval(result,{}, {"workbook":self.sheet.workbook})
                return self.val
            except:
                if strict:
                    raise exc_info()[0]("\nBad Formula Expression: %s\n\n\tEval:\t%s\n\tOrig:\t%s"%(str(exc_info()[1]), result, self.fn))
                else:
                    return False
        def __setattr__(self, key, val):
            if key == "fn":
                pass
            self.__dict__[key]=val
        
        def __int__(self):
            return int(self.val)
        
        def __float__(self):
            return float(self.val)
        
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


if __name__ == "__main__":
    Workbook = workbook("test.xlsx")