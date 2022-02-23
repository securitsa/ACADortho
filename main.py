from tkinter import *
from pyautocad import *
from math import *
import win32com.client
import copy


def main():
    window = Tk()
    my_gui = GUI(window)
    window.mainloop()


def search_line():
    acad = Autocad(create_if_not_exists=True)
    print(acad.doc.name)
    line = []
    i = 0

    for item in acad.iter_objects_fast('Line'):
        if item.ObjectName == 'AcDbLine':
            line.append(Line())
            line[i].start_point = list(item.StartPoint)
            line[i].end_point = list(item.EndPoint)
            Line.get_position(line[i])
            i += 1

    poly_coord = []
    poly_line = []

    for item in acad.iter_objects_fast("PolyLine"):
        poly_coord = item.Coordinates

    poly_coord = list(poly_coord)

    if poly_coord:
        poly_line.append(Line())
        poly_point = []
        poly_point.append(poly_coord[0])
        poly_point.append(poly_coord[1])
        poly_point.append(0.0)

        poly_line[0].start_point = copy.copy(poly_point)
        end_point_x = poly_coord[-2]
        end_point_y = poly_coord[-1]

        poly_coord.pop(0)
        poly_coord.pop(0)
        poly_coord.pop(-1)
        poly_coord.pop(-1)

    # poly_coord.pop(len(poly_coord)-1)
    # poly_coord.pop(len(poly_coord)-2)

        poly_point.clear()
        i = 0

        for item in poly_coord:
            poly_point.append(item)
            if len(poly_point) == 2:
                poly_point.append(0.0)
                poly_line[i].end_point = copy.copy(poly_point)
                poly_line.append(Line())
                i += 1
                poly_line[i].start_point = copy.copy(poly_point)
                poly_point.clear()
            else:
                continue

        poly_point.append(end_point_x)
        poly_point.append(end_point_y)
        poly_point.append(0.0)

        poly_line[i].end_point = copy.copy(poly_point)
        poly_point.clear()

    # for item in poly_line:
    #     print(item.start_point, item.end_point)

    for i in poly_line:
        line.append(i)

    for item in line:
        Line.edit_coordinates(item)
        print(item.start_point, item.end_point)

    num_line = len(line)
    poly = len(poly_line)
    edit = 0
    for item in line:
        if item.type == 3:
            edit += 1



    acad1 = win32com.client.Dispatch("AutoCAD.Application")
    acad1.Visible = True
    acad1Model = acad1.ActiveDocument.ModelSpace

    for item in acad1Model:
        item.Delete()

    for item in line:
        p1 = APoint(item.start_point)
        p2 = APoint(item.end_point)
        acad.model.AddLine(p1, p2)



class Line:
    def __init__(self):
        self.start_point = []
        self.end_point = []
        self.deviation = [0, 0, 0]
        self.type = int

    def get_position(self):
        if abs(self.start_point[0] - self.end_point[0]) <= self.deviation[0]:
            self.type = 0
            return 0
        elif abs(self.start_point[1] - self.end_point[1]) <= self.deviation[1]:
            self.type = 1
            return 1
        elif abs(self.start_point[2] - self.end_point[2]) <= self.deviation[2] and self.start_point[2] != 0:
            self.type = 2
            return 2
        else:
            self.type = 3
            return 3

    def add_start_point(self, start_point):
        self.start_point = start_point

    def add_end_point(self, end_point):
        self.end_point = end_point


    def get_position_type(self):
        if self.get_position() == 0:
            return str("Vertical")

        elif self.get_position() == 1:
            return str("Horizontal")

        elif self.get_position() == 2:
            return str("Parallel z")

        else:
            return str("Low position")

    def get_type(self):
        return self.type

    def get_tan(self):
        if self.type == 3:
            tan = (abs(self.end_point[1]-self.start_point[1])) / (abs(self.end_point[0] - self.start_point[0]))
            return tan

    def edit_coordinates(self):
        if self.type == 3:
            if self.get_tan() > 1:
                start_p = self.start_point[0] + (self.end_point[0]-self.start_point[0])/2
                end_p = self.end_point[0] - (self.end_point[0]-self.start_point[0])/2
                self.end_point.pop(0)
                self.end_point.insert(0, int(end_p))
                self.start_point.pop(0)
                self.start_point.insert(0, int(start_p))

            elif self.get_tan() < 1 and self.get_tan() > 0:
                start_p_1 = self.start_point[1] + (self.end_point[1] - self.start_point[1])/2
                end_p_1 = self.end_point[1] - (self.end_point[1] - self.start_point[1])/2
                self.end_point.pop(1)
                self.end_point.insert(1, end_p_1)
                self.start_point.pop(1)
                self.start_point.insert(1, start_p_1)

class GUI:
    line = int
    poly_line = int
    edit_line = int

    def __init__(self, window) -> None:
        self.dX = 1
        self.dY = 1
        self.dZ = 1
        self.window = window
        window.title('Зубчатое колесо')
        window.geometry('640x280')
        self.selected = IntVar()
        self.paramLable = Label(
            window, text="Параметры колеса", font=("Arial Bold", 14))
        self.dX_text = Label(window, text="dX",
                            font=("Arial Bold", 10))

        self.dX_text.grid(column=0, row=10)
        self.dY_text = Label(window, text="dY",
                            font=("Arial Bold", 10))
        self.dY_text.grid(column=0, row=11)
        self.dZ_text = Label(window, text="dZ",
                                 font=("Arial Bold", 10))
        self.dZ_text.grid(column=0, row=12)
        self.dZ_enter = Entry(window, width=10)
        self.dZ_enter.insert(END, '1')
        self.dZ_enter.grid(column=1, row=12)
        self.dX_enter = Entry(window, width=10)
        self.dX_enter.insert(END, '1')
        self.dX_enter.grid(column=1, row=10)
        self.dY_enter = Entry(window, width=10)
        self.dY_enter.insert(END, '1')
        self.dY_enter.grid(column=1, row=11)
        self.ok_button = Button(text='Ok', command=self.collectData, width=15)
        self.ok_button.grid(column=0, row=13)
        self.cancel_button = Button(
            text='Cancel', command=self.quit, width=15)
        self.cancel_button.grid(column=1, row=13)




    def quit(self):
        self.window.destroy()


    def info(self, window, line, edit_line, poly_line):
        self.dX_text = Label(window, text="Найдено",
                             font=("Arial Bold", 10))
        self.dX_text.grid(column=4, row=10)

        self.dX_text = Label(window, text=line,
                             font=("Arial Bold", 10))
        self.dX_text.grid(column=6, row=10)




    def collectData(self):
        self.dX = float(self.dX_enter.get())
        self.dY = float(self.dY_enter.get())
        self.dZ = float(self.dZ_enter.get())
        search_line()


    def getData(self):
        return self.dX, self.dY, self.dZ


if __name__ == "__main__":
    main()