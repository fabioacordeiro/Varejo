# deepcopys almost everything (functions, lambda ... ) 

## pip install deepcopyall

```python

from deepcopyall import deepcopy
class bibi:
    def __init__(self, bebe=112):
        self.bebe = bebe

    def printbebe(self):
        print(self.bebe)


baba = lambda: print("baba")


def buax(x):
    print(x)
    return x


baba2 = deepcopy(o=baba)
baba2()
buax2 = deepcopy(o=buax)
q = buax2(444)
print(q)

bibi2 = deepcopy(o=bibi)
bibi2inst = bibi2(23323)
bibi2inst.printbebe()

#output 
baba
444
444
23323


```