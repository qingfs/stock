# -*- coding: utf-8 -*-
class Duck(object):
    def run(self):
        print("runing...")

    @property
    def name(self):
        return self.__name

    @name.setter
    def name(self, value):
        self.__name = value


class Time(object):
    def run(self):
        print("runing...")


class Student(object):

    @property
    def score(self):
        return self._score

    @score.setter
    def score(self, value):
        if not isinstance(value, int):
            raise ValueError('score must be an integer!')
        if value < 0 or value > 100:
            raise ValueError('score must between 0 ~ 100!')
        self._score = value


if __name__ == '__main__':
    d = Duck()
    d.name = 'text'
    print(d.name)

