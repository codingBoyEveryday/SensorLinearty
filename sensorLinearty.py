#!/usr/bin/python
import xlrd
import matplotlib.pyplot as plt
import math

from numpy import *


class SensorLinearty():

    def __init__(self):

        self.required_data_time = []
        self.required_data_angle = []
        self.new_required_data_time = []
        self.new_required_data_angle = []
        self.startPoint = 0


    def open_data_file(self, filePath):

        workbook = xlrd.open_workbook(filePath)
        worksheet = workbook.sheet_by_index(0)
        return worksheet


    def get_data(self, current_worksheet, column_1, column_2):
        """if we display the time by second, then the figure is not so good
        #time_tuple = xlrd.xldate_as_tuple(exceltime,0)
        #time_in_second = time_tuple[4]*60 + time_tuple[5]
        #required_data_time.append(time_in_second)
        """

        for rownum in range(current_worksheet.nrows):
            row_values = current_worksheet.row_values(rownum)           ### get values in all columns of every row
            exceltime = current_worksheet.cell_value(rownum,column_1)   ### get the cell value in the column column_1
            self.required_data_time.append(exceltime)                   ### make the list storing time
            self.required_data_angle.append(row_values[column_2])       ### make the list storing degree using data of column_2


        for i in range(len(self.required_data_angle)):
            if (self.required_data_angle[i+1]-self.required_data_angle[i])>=2:
                self.startPoint = i
                break

        self.new_required_data_time[:] = self.required_data_time[self.startPoint:]
        # print self.new_required_data_time
        self.new_required_data_angle[:] = self.required_data_angle[self.startPoint:]
        # print self.new_required_data_angle



    def ajust_angle(self):

        for i in range(len(self.new_required_data_angle)):
            if (self.new_required_data_angle[i]-self.new_required_data_angle[i-1]) > 250:
                self.new_required_data_angle[i] = self.new_required_data_angle[i] - 360
            elif (self.new_required_data_angle[i]-self.new_required_data_angle[i-1]) < -250:
                self.new_required_data_angle[i] = self.new_required_data_angle[i] + 360




    def plot_figure(self):

        plt.plot(self.new_required_data_time, self.new_required_data_angle)
        plt.show()








def main():
    testPathFile = "C:/Users/HPuser/Desktop/sensorLinearty/test2.xlsx"

    example_1 = SensorLinearty()
    example_2 = SensorLinearty()
    example_3 = SensorLinearty()
    example_4 = SensorLinearty()
    example_5 = SensorLinearty()

    testWorksheet_1 = example_1.open_data_file(testPathFile)
    testWorksheet_2 = example_2.open_data_file(testPathFile)
    testWorksheet_3 = example_3.open_data_file(testPathFile)
    testWorksheet_4 = example_4.open_data_file(testPathFile)
    testWorksheet_5 = example_5.open_data_file(testPathFile)



    example_1.get_data(testWorksheet_1, 1, 3)
    example_1.ajust_angle()
    # example_1.plot_figure()

    example_2.get_data(testWorksheet_2, 5, 7)
    example_2.ajust_angle()
    # example_2.plot_figure()

    example_3.get_data(testWorksheet_3, 9, 11)
    example_3.ajust_angle()
    # example_3.plot_figure()

    example_4.get_data(testWorksheet_4, 13, 15)
    example_4.ajust_angle()
    # example_4.plot_figure()

    example_5.get_data(testWorksheet_5, 17, 19)
    example_5.ajust_angle()
    # example_5.plot_figure()


######### ajust the start point of time #################
    min_start_time_point = min(example_1.new_required_data_time[0], example_2.new_required_data_time[0],
                               example_3.new_required_data_time[0], example_4.new_required_data_time[0],
                               example_5.new_required_data_time[0],)
    t1 = example_1.new_required_data_time[0] - min_start_time_point
    t2 = example_2.new_required_data_time[0] - min_start_time_point
    t3 = example_3.new_required_data_time[0] - min_start_time_point
    t4 = example_4.new_required_data_time[0] - min_start_time_point
    t5 = example_5.new_required_data_time[0] - min_start_time_point

    for i in range(len(example_1.new_required_data_time)):
        example_1.new_required_data_time[i] = example_1.new_required_data_time[i] - t1
        example_2.new_required_data_time[i] = example_2.new_required_data_time[i] - t2
        example_3.new_required_data_time[i] = example_3.new_required_data_time[i] - t3
        example_4.new_required_data_time[i] = example_4.new_required_data_time[i] - t4
        example_5.new_required_data_time[i] = example_5.new_required_data_time[i] - t5

#########################################################

#

####### adjust the start point in angle #################

    # min_start_angle_point = min(example_1.new_required_data_angle[0], example_2.new_required_data_angle[0],
    #                            example_3.new_required_data_angle[0], example_4.new_required_data_angle[0],
    #                            example_5.new_required_data_angle[0], )
    # a1 = example_1.new_required_data_angle[0] - min_start_angle_point
    # a2 = example_2.new_required_data_angle[0] - min_start_angle_point
    # a3 = example_3.new_required_data_angle[0] - min_start_angle_point
    # a4 = example_4.new_required_data_angle[0] - min_start_angle_point
    # a5 = example_5.new_required_data_angle[0] - min_start_angle_point
    #
    # for i in range(len(example_1.new_required_data_angle)):
    #     example_1.new_required_data_angle[i] = example_1.new_required_data_angle[i] - a1
    #     example_2.new_required_data_angle[i] = example_2.new_required_data_angle[i] - a2
    #     example_3.new_required_data_angle[i] = example_3.new_required_data_angle[i] - a3
    #     example_4.new_required_data_angle[i] = example_4.new_required_data_angle[i] - a4
    #     example_5.new_required_data_angle[i] = example_5.new_required_data_angle[i] - a5





################################################################

    plt.plot(example_1.new_required_data_time, example_1.new_required_data_angle, 'r')
    plt.plot(example_2.new_required_data_time, example_2.new_required_data_angle, 'g')
    plt.plot(example_3.new_required_data_time, example_3.new_required_data_angle, 'b')
    plt.plot(example_4.new_required_data_time, example_4.new_required_data_angle, 'y')
    plt.plot(example_5.new_required_data_time, example_5.new_required_data_angle, 'm')

    plt.show()

if __name__ == "__main__":
    main()

