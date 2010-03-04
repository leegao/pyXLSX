import re

#Interpolation Test

def transpose(t):
    a = re.compile(".")
    _t = a.findall(t)
    n = 0
    for i in range(len(t)):
        n+= 26**(len(t)-1-i)*(ord(t[i])-64)
    return n-1

def range_alpha(start, end):
    #Case 1: A-Z
    if len(start)==1 and len(end) == 1:
        return [chr(n) for n in range(ord(start), ord(end)+1)]
    #Case 2: A-AA
    if len(start) <= len(end):
        _r = _ra(len(end))
        return _r[transpose(start):transpose(end)+1]
    
def _ra(n, l=range_alpha("A","Z"), f = range_alpha("A","Z")):
    _l= []
    n-=1
    if n > 0:
        for c in l:
            _l += [c+a for a in range_alpha("A","Z")]
        f += _l  
        return _ra(n, _l, f)
    else:
        return f
       
def PARSE(match):
    return "worksheet."+match.group().upper()

def REPLACE(match):
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
        return str(_range)
    
    #Case 2: Same numer
    if N0 == N1:
        #Range in alpha
        _range = [a+str(N0) for a in range_alpha(A0, A1)]
        return str(_range)
    
    #Case 3: Else
    
    ###Unimplemented Yet###
    
    return tokens

def interpolate(text):
    rn = re.compile(r"[a-z|A-Z]+[0-9]+\:*[a-z|A-Z]+[0-9]+\:*")
    fn = re.compile(r"(?<=\(|\))[a-z|A-Z]*(?=\()")
    print rn.sub(REPLACE, fn.sub(PARSE, text))

interpolate("6+(SUM(Z1:AH1)+3)")
#6+(worksheet.SUM(['Z1', 'AA1', 'AB1', 'AC1', 'AD1', 'AE1', 'AF1', 'AG1', 'AH1'])+3)