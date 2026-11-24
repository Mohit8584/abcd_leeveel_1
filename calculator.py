import time

class Calculator:
    def __init__(self):
        print('welcome and do add substraction')
        self.__history=[]

    def add(self,*args):
        total=0
        for i in args:
           total += i
        t=time.time()
        self.__history.append({f"time taken -:-{t}" : {args : total }})

        return total     

    def get_history(self):
        print(self.__history)    

def subs(a,b):
    if a>b:
        return a-b
    else:
        return b-a


print(__name__)

if __name__=='__main__':
    c=subs(3,1)
    print(c)

