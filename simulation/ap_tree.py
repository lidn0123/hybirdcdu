import os
import win32com.client as win32

aspen = win32.Dispatch('Apwn.Document')
print('generate aspen+ com')
aspen.InitFromArchive2(os.path.join(os.path.dirname(__file__), 'hbcdumodel/CDU-basic.bkp'))
print('open file finish')

f = open('tree.txt', 'x')


def print_elements(obj, level=0):
    if hasattr(obj, 'Elements'):
        print(' ' * level + obj.Name)
        f.write(' ' * level + obj.Name)
        for o in obj.Elements:
            print_elements(o, level + 1)
    else:
        print(' ' * level, obj.Name, ' = ', obj.Value)
        f.write(' ' * level, obj.Name, ' = ', obj.Value)


print_elements(aspen.Tree)
f.close()

aspen.Close
