import re

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