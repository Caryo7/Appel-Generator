from graph import *

class Progress:
    def __init__(self, title, length, larg):
        self.lines = [setLines(title), emptyLine()]
        self.length = length
        self.pos = 0
        self.values = []
        self.color_bar = 'black'
        self.larg = larg + 2
        self.update()

    def format_line(self, text):
        while len(text) < self.larg:
            text += ' '

        return text

    def step(self, name = None, color = None, bar = None):
        self.pos += 1
        if name is not None:
            name = self.format_line(name)
            if color is not None:
                d = c(name, style = {'fg': color})
            else:
                d = name

            self.values.append(setLines(d))

        if bar is not None:
            self.color_bar = bar

        self.update()

    def getFiveLast(self):
        lst = self.values[-5:][::-1]
        while len(lst) < 5:
            lst.append(setLines(' '*self.larg))

        return lst

    def update(self):
        pc = setLines(c(progress(self.pos/self.length), style = {'fg': self.color_bar}))
        lns = self.lines + [pc, emptyLine()] + self.getFiveLast()
        text = centerText(*lns)
        finalPrint(text, fnct = None)
        self.color_bar = 'black'
