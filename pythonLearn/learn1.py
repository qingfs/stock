# -*- coding: utf-8 -*-
from docx import Document
def hello(name):
    strHello = 'Hello, ' + name
    return strHello;


def learnList():
    classmates = ('a')
    print(classmates)


def add(x, y, f):
    return f(x) + f(y)


if __name__ == '__main__':
    d1 = Document()
    p1 =d1.add_paragraph()
    for i in range(1, 4):
        print(i)

