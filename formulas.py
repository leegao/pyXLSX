from __future__ import division
from xlsx import workbook, _flatten
import math, cmath, fractions, random

extend = workbook.extend #Decorator for adding new formulas
flatten = workbook.flatten #Formula doesn't care if a list is one or multi dimensional
singular = workbook.singular #Only takes in a single value
plural = workbook.plural #Takes multiple arguments, does not check for first param
void = workbook.void #Takes no arguments

def integrate(fn, lower, upper, threshold):
    dx = float(upper-lower)/threshold
    ddx = dx + lower
    fn_dx = 0
    while ddx < upper:
        fn_dx += fn(ddx)*dx
        ddx += dx
    return fn_dx

def im(z):
    return complex(str(z).replace("i", "j"))

def _round(x, sig, key=round):
    sign = 1 if x>0 else -1
    sig = -math.log10(abs(sig))
    x = abs(x)
    return sign*float(key(x*(10**sig)))/(10**sig)

def _roman(input):
    if type(input) != type(1):
        raise TypeError, "expected integer, got %s" % type(input)
    if not 0 < input < 4000:
        raise ValueError, "Argument must be between 1 and 3999"   
    ints = (1000, 900,  500, 400, 100,  90, 50,  40, 10,  9,   5,  4,   1)
    nums = ('M',  'CM', 'D', 'CD','C', 'XC','L','XL','X','IX','V','IV','I')
    result = ""
    for i in range(len(ints)):
       count = int(input / ints[i])
       result += nums[i] * count
       input -= ints[i] * count
    return result

@extend
@flatten
def SUM(list):
    "Sums up the list of numbers"
    sigma = 0
    for n in list:
        sigma += n
    return sigma

@extend
@flatten
def AVERAGE(list):
    "Find the average of the list of numbers"
    sigma = 0.0
    i = 0
    for n in list:
        sigma += n
        i+=1
    return sigma/i

@extend
@singular
def CODE(char):
    "Return the first char code"
    return ord(str(char)[0])

@extend
@plural
def BESSELI(base, nu):
    """
    Returns the Modified BESSEL of First Kind function of parameters base and nu respectively.
    BESSELI ~ (z,v): 1/pi integral_0^pie^(zcos(t)) cos(vt) dt
    """
    dfn_dx = lambda x: (math.e**(base*math.cos(x)))*math.cos(nu*x)
    return (1./math.pi)*integrate(dfn_dx, 0, math.pi, 5000)

@extend
@plural
def BESSELJ(base, nu):
    """
    Returns the BESSEL Function of First Kind of parameters base and nu respectively.
    BESSELJ ~ (z,v): 1/pi integral_0^pie^cos(zsin(t) - vt) dt
    """
    dfn_dx = lambda x: math.cos(base*math.sin(x)-nu*x)
    return (1./math.pi)*integrate(dfn_dx, 0, math.pi, 5000)

@extend
@plural
def BESSELK(base, nu):
    """
    Returns the BESSEL Function of Second Kind of parameters base and nu respectively.
    BESSELJ ~ (z,v): pi/2 (I-z(v)-Iz(v))/sin(z*pi)
    
    BROKEN IMPLEMENTATION
    """
    I = lambda n,x: (1./math.pi)*integrate(lambda x: (math.e**(base*math.cos(x)))*math.cos(nu*x), 0, math.pi, 5000)
    #print 2*I(-1.5, 1)/(math.sin(base*math.pi))
    K = (math.pi/2)*(I(-base, nu)-I(base, nu))/(math.sin(base*math.pi))
    return K

@extend
@singular
def BIN2DEC(binary):
    bin = list(str(binary))
    dec = 0
    for i in range(len(bin)):
        dec += int(bin[::-1][i])*2**i
    return dec
    
@extend
@singular
def DEC2BIN(dec):
    return bin(dec).replace("0b","")

@extend
@singular
def BIN2HEX(binary):
    bin = list(str(binary))
    dec = 0
    for i in range(len(bin)):
        dec += int(bin[::-1][i])*2**i
    return "%X" % dec

@extend
@singular
def DEC2HEX(dec):
    return "%X"%dec

@extend
@singular
def BIN2OCT(binary):
    bin = list(str(binary))
    dec = 0
    for i in range(len(bin)):
        dec += int(bin[::-1][i])*2**i
    return "%o" % dec

@extend
@singular
def DEC2OCT(dec):
    return "%o"%dec
    
@extend
@singular
def HEX2DEC(dec):
    return int(dec, 16)
    
@extend
@singular
def HEX2OCT(dec):
    return "%o"%int(dec, 16)
    
@extend
@singular
def HEX2BIN(dec):
    return bin(dec).replace("0b","")
    
@extend
@singular
def OCT2DEC(dec):
    return int(dec, 8)
    
@extend
@singular
def OCT2HEX(dec):
    return "%X"%int(dec, 8)
    
@extend
@singular
def OCT2BIN(dec):
    return "%b"%int(dec, 8)

@extend
@plural
def COMPLEX(real, imaginary):
    return complex(real, imaginary)

@extend
@plural
def DELTA(a, b):
    return 1 if a == b else 0

@extend
@plural
def ERF(a, b = 0):
    """http://office.microsoft.com/en-us/excel/HP052090771033.aspx"""
    def _erf(z):
        return (2./math.sqrt(math.pi))*integrate(lambda x: math.e**(-x**2), 0, z, 5000)
    if not b:
        return _erf(a)
    else:
        return _erf(b) - _erf(a)

@extend
@singular
def ERFC(z):
    def _erfc(z):
        return (2./math.sqrt(math.pi))*integrate(lambda x: math.e**(-x**2), 0, z, 5000)
    return 1 - _erfc(z)

@extend
@plural
def GESTEP(a,b):
    return 1 if a >= b else 0

@extend
@singular
def IMABS(z):
    c = im(z)
    return abs(c)

@extend
@singular
def IMAGINARY(z):
    c = im(z)
    return c.imag

@extend
@singular
def IMARGUMENT(z):
    c = im(z)
    return cmath.phase(c)

@extend
@singular
def IMCONJUGATE(z):
    c = im(z)
    return c.conjugate()

@extend
@singular
def IMCOS(z):
    c = im(z)
    return cmath.cos(c)

@extend
@plural
def IMDIV(a,b):
    a = im(a)
    b = im(b)
    return a/b

@extend
@singular
def IMEXP(z):
    c = im(z)
    return cmath.exp(c)

@extend
@singular
def IMLN(z):
    c = im(z)
    return cmath.log(c)

@extend
@singular
def IMLOG10(z):
    c = im(z)
    return cmath.log10(c)

@extend
@singular
def IMLOG2(z):
    c = im(z)
    return cmath.log(c, 2)

@extend
@plural
def IMPOWER(a,b):
    a = im(a)
    return a**b

@extend
@plural
def IMPRODUCT(a,b):
    a = im(a)
    b = im(b)
    return a*b

@extend
@singular
def IMREAL(z):
    c = im(z)
    return c.real

@extend
@singular
def IMSIN(z):
    c = im(z)
    return cmath.sin(c)

@extend
@singular
def IMSQRT(z):
    c = im(z)
    return cmath.sqrt(c)

@extend
@plural
def IMSUB(a,b):
    a = im(a)
    b = im(b)
    return a-b

@extend
@plural
def IMSUM(a,b):
    a = im(a)
    b = im(b)
    return a+b


######################################

@extend
@singular
def ABS(x):
    return abs(x)

@extend
@singular
def ACOS(x):
    return math.acos(x)

@extend
@singular
def ACOSH(x):
    return math.acosh(x)

@extend
@singular
def ASIN(x):
    return math.asin(x)

@extend
@singular
def ASINH(x):
    return math.asinh(x)

@extend
@singular
def ATAN(x):
    return math.atan(x)

@extend
@singular
def ATANH(x):
    return math.atanh(x)

@extend
@plural
def ATAN2(x, y):
    return math.atan2(x, y)

@extend
@plural
def CEILING(x, sig):
    sign = 1 if x>0 else -1
    sig = -math.log10(abs(sig))
    x = abs(x)
    return float(math.ceil(x*(10**sig)))/(10**sig)

@extend
@plural
def COMBIN(n, k):
    def _prod(fn, lower, upper, n = 1):
        for i in range(lower, upper+1):n *= fn(i)
        return n
    x = lambda x: x if x else 1
    return _prod(x, k+1, n)/_prod(x, 1, n-k)

@extend
@singular
def COS(x):
    return math.cos(x)

@extend
@singular
def COSH(x):
    return math.cosh(x)

@extend
@singular
def DEGREES(theta):
    return theta/math.pi*180

@extend
@singular
def EVEN(x):
    sign = 1 if x>0 else -1
    if not _round(x, sign)%2: return int(_round(x, sign))
    x = abs(x)
    return sign*int(math.ceil(float(x)/2)*2)

@extend
@singular
def EXP(x):
    return math.exp(x)

@extend
@singular
def FACT(n):
    def _prod(fn, lower, upper, n = 1):
        for i in range(int(lower), int(upper+1)):n *= fn(i)
        return n
    x = lambda x: x if x else 1
    return _prod(x, 1, n)

@extend
@singular
def FACTDOUBLE(n):
    x = 1
    while n > 0:
        x *= n
        n -= 2
    return x

@extend
@plural
def FLOOR(x, sig):
    sign = 1 if x>0 else -1
    sig = -math.log10(abs(sig))
    x = abs(x)
    return sign*float(math.floor(x*(10**sig)))/(10**sig)

@extend
@plural
def GCD(x, y):
    return fractions.gcd(x, y)

@extend
@singular
def INT(x):
    return int(math.floor(x))

@extend
@plural
def LCM(x, y):
    return (x*y)/fractions.gcd(x, y)

@extend
@singular
def LN(x):
    return math.log(x)

@extend
@plural
def LOG(x, y):
    return math.log(x, y)

@extend
@singular
def LOG10(x):
    return math.log10(x)

@extend
def MDETERM(matrix):
    if len(matrix) is not len(matrix[0]): return
    matrix = [[float(e) for e in m] for m in matrix]
    def _det(matrix):
        def _minor(matrix, i, j):
            _matrix = [[e for e in m] for m in matrix]
            _matrix.pop(i)
            for _m in _matrix: _m.pop(j)
            return _matrix
        n = 0
        if len(matrix) == 2:
            return matrix[0][0]*matrix[1][1] - matrix[0][1]*matrix[1][0] #ad-bc
        for j in range(len(matrix)):
            n += matrix[0][j]*(-1)**j*_det(_minor(matrix,0,j))
        return n
    
    return _det(matrix)

@extend
def MINVERSE(matrix):
    if len(matrix) is not len(matrix[0]): return
    matrix = [[float(e) for e in m] for m in matrix]
    def _minor(matrix, i, j):
        _matrix = [[e for e in m] for m in matrix]
        _matrix.pop(i)
        for _m in _matrix: _m.pop(j)
        return _matrix
    def _det(matrix):
        n = 0
        if len(matrix) == 2:
            return matrix[0][0]*matrix[1][1] - matrix[0][1]*matrix[1][0] #ad-bc
        for j in range(len(matrix)):
            n += matrix[0][j]*(-1)**j*_det(_minor(matrix,0,j))
        return n
    
    A_inv = 1./_det(matrix)
    inverse = [[0 for e in m] for m in matrix]
    if len(matrix) == 2:
        matrix[0][0],matrix[1][1] = -matrix[1][1],-matrix[0][0]
        return [[-A_inv*float(e) for e in m] for m in matrix]
    for i in range(len(matrix)):
        for j in range(len(matrix)):
            inverse[j][i] = A_inv*(-1)**(i+j)*_det(_minor(matrix, i, j))
    return inverse

@extend
@plural
def MMULT(matrix1, matrix2):
    if len(matrix1[0]) is not len(matrix2[0]): return
    _m = []
    for row1 in matrix1:
        _r = []
        for col2 in zip(*matrix2):
            _r += [sum([row1[i]*col2[i] for i in range(len(row1))])]
        _m += [_r]
    return _m

@extend
@plural
def MOD(a,b):
    return a%b

@extend
@plural
def MROUND(a,b):
    n = min(a%b, a%b-b, key = lambda a: abs(a)) if int(2*a%b - b) else a%b-b
    return a-n

@extend
@plural
def MULTINOMIAL(*all):
    def _fact(n):
        return n*_fact(n-1) if n else 1
    n = 1
    for i in all:
        n *= _fact(i)
    return float(_fact(sum(all)))/n

@extend
@singular
def ODD(x):
    sign = 1 if x>0 else -1
    if _round(x, sign)%2: return int(_round(x, sign))
    x = abs(x)
    return sign*int(math.ceil(float(x)/2)*2 + 1)

@extend
@void
def PI(): return math.pi

@extend
@plural
def POWER(a,b):
    return a**b

@extend
@flatten
def PRODUCT(list):
    n = 1
    for i in list:
        n*=i
    return n

@extend
@plural
def QUOTIENT(a,b):
    return int(_round(a/b, 1, key=math.floor))

@extend
@singular
def RADIANS(theta):
    return math.pi*theta/180

@extend
@void
def RAND():
    return random.random()

@extend
@plural
def RANDBETWEEN(a,b):
    return random.randint(a,b)

@extend
@plural
def ROMAN(a,b):
    del b
    return _roman(a)

@extend
@plural
def ROUND(n,sig):
    return _round(n, 10**-sig)

@extend
@plural
def ROUNDDOWN(n,sig):
    return _round(n, 10**-sig, key=math.floor)

@extend
@plural
def ROUNDUP(n,sig):
    return _round(n, 10**-sig, key=math.ceil)

@extend
@plural
def SERIESSUM(x, n, m, a):
    a = _flatten(a)
    r = 0
    for i in range(len(a)):
        r += a[i]*(x**(n+m*i))
    return r

@extend
@singular
def SIGN(x):
    return 0 if not x else 1 if x>0 else -1

@extend
@singular
def SIN(x):
    return math.sin(x)

@extend
@singular
def SINH(x):
    return math.sinh(x)

@extend
@singular
def SQRT(x):
    return math.sqrt(x)

@extend
@singular
def SQRTPI(x):
    return math.sqrt(math.pi*x)

@extend
@flatten
def SUBTOTAL(all):
    fn = all.pop(0)
    fn_dict = {
               1:workbook.AVERAGE,
               2:workbook.COUNT,
               3:workbook.COUNTA,
               4:workbook.MAX,
               5:workbook.MIN,
               6:workbook.PRODUCT,
               7:workbook.STDEV,
               8:workbook.STDEVP,
               9:workbook.SUM,
               10:workbook.VAR,
               11:workbook.VARP
               }
    return fn_dict[int(fn)](all)

@extend
@flatten
def SUM(list):
    n = 0
    for i in list:
        n+=float(i)
    return n

@extend
@plural
def SUMIF(a, cond, b = []):
    a,b = _flatten(a),_flatten(b)
    if not b: b = a
    n=  0
    o_cond = cond
    if "<" not in str(cond) and ">" not in str(cond): cond = " == "+str(cond)
    for i in range(len(a)):
        try:
            if eval(str(a[i]) + cond):
                n+= float(b[i])
        except:
            if eval(str(a[i]) + "== '"+o_cond+"'"):
                n+= float(b[i])
    return n

@extend
def SUMPRODUCT(*list):
    n = 0
    for i in range(len(list)):
        _n = 1
        for l in list:
            _n*=_flatten(l)[i]
        n+=_n
    return n

if __name__ == "__main__":
#    I = lambda b,n: (1./cmath.pi)*integrate(lambda x: (cmath.e**(b*cmath.cos(x)))*cmath.cos(n*x), 0, cmath.pi, 5000)
#    
#    J = lambda b,n: (1./cmath.pi)*integrate(lambda x: cmath.cos(b*cmath.sin(x)-n*x), 0, cmath.pi, 5000)
#    
#    i = complex(0,1.)
#    z = 2.5
#    v = 1
#    print math.sin(v*math.pi)
    import demo
    pass