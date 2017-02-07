#!/usr/bin/python
import xlrd
import matplotlib.pyplot as plt
import math
import numpy as np
import pylab

# from numpy import *


class SensorLinearty():

    def __init__(self):

        self.required_data_time = []
        self.required_data_angle = []
        self.new_required_data_time = []
        self.new_required_data_angle = []
        self.matrix = []
        self.startPoint = 0
        self.endPoint = 0

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

    def adjust_start_time(self):
        for i in range(len(self.required_data_angle)):
            if abs((self.required_data_angle[i+1]-self.required_data_angle[i]))>=3:
                self.startPoint = i
                break
        self.new_required_data_time[:] = self.required_data_time[self.startPoint:]
        # print self.new_required_data_time
        self.new_required_data_angle[:] = self.required_data_angle[self.startPoint:]
        # print self.new_required_data_angle

    # def remove_useless_timerange(self):
    #     for i in range(len(self.new_required_data_angle)):
    #         if abs((self.required_data_angle[i+1]-self.required_data_angle[i]))==0:
    #             self.endPoint = i
    #             break
    #     self.new_required_data_time[:] = self.required_data_time[0:self.endPoint]
        # print self.new_required_data_time
        # self.new_required_data_angle[:] = self.required_data_angle[0:self.endPoint]
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


    def acc_rebuild(self):
        pass








def main():
    mechineSensor_testPathFile = "C:/Users/HPuser/Desktop/sensorLinearty/machine.xlsx"
    accSensor_testPathFile = "C:/Users/HPuser/Desktop/sensorLinearty/acc.xlsx"

    example_1 = SensorLinearty()
    example_2 = SensorLinearty()
    example_3 = SensorLinearty()
    example_4 = SensorLinearty()
    example_5 = SensorLinearty()
    acc = SensorLinearty()

    testWorksheet_1 = example_1.open_data_file(mechineSensor_testPathFile)
    testWorksheet_2 = example_2.open_data_file(mechineSensor_testPathFile)
    testWorksheet_3 = example_3.open_data_file(mechineSensor_testPathFile)
    testWorksheet_4 = example_4.open_data_file(mechineSensor_testPathFile)
    testWorksheet_5 = example_5.open_data_file(mechineSensor_testPathFile)
    testWorksheet_acc = acc.open_data_file(accSensor_testPathFile)



    example_1.get_data(testWorksheet_1, 1, 3)
    example_1.adjust_start_time()
    example_1.ajust_angle()
    # example_1.plot_figure()

    example_2.get_data(testWorksheet_2, 5, 7)
    example_2.adjust_start_time()
    example_2.ajust_angle()
    # example_2.plot_figure()

    example_3.get_data(testWorksheet_3, 9, 11)
    example_3.adjust_start_time()
    example_3.ajust_angle()
    # example_3.plot_figure()

    example_4.get_data(testWorksheet_4, 13, 15)
    example_4.adjust_start_time()
    example_4.ajust_angle()
    # example_4.plot_figure()

    example_5.get_data(testWorksheet_5, 17, 19)
    example_5.adjust_start_time()
    example_5.ajust_angle()
    # example_5.plot_figure()


    acc.get_data(testWorksheet_acc,8, 14)
    acc.adjust_start_time()
    acc.ajust_angle()
    # acc.plot_figure()


######### ajust the start point of time #################
    min_start_time_point = min(example_1.new_required_data_time[0], example_2.new_required_data_time[0],
                               example_3.new_required_data_time[0], example_4.new_required_data_time[0],
                               example_5.new_required_data_time[0], acc.new_required_data_time[0])
    # print min_start_time_point

    t1 = example_1.new_required_data_time[0] - min_start_time_point
    t2 = example_2.new_required_data_time[0] - min_start_time_point
    t3 = example_3.new_required_data_time[0] - min_start_time_point
    t4 = example_4.new_required_data_time[0] - min_start_time_point
    t5 = example_5.new_required_data_time[0] - min_start_time_point
    t_acc = acc.new_required_data_time[0] - min_start_time_point

    for i in range(len(example_1.new_required_data_time)):
        example_1.new_required_data_time[i] = example_1.new_required_data_time[i] - t1
    for i in range(len(example_2.new_required_data_time)):
        example_2.new_required_data_time[i] = example_2.new_required_data_time[i] - t2
    for i in range(len(example_3.new_required_data_time)):
        example_3.new_required_data_time[i] = example_3.new_required_data_time[i] - t3
    for i in range(len(example_4.new_required_data_time)):
        example_4.new_required_data_time[i] = example_4.new_required_data_time[i] - t4
    for i in range(len(example_5.new_required_data_time)):
        example_5.new_required_data_time[i] = example_5.new_required_data_time[i] - t5
    for i in range(len(acc.new_required_data_time)):
        acc.new_required_data_time[i] = acc.new_required_data_time[i] - t_acc

#########################################################

#

####### adjust the start point in angle #################

    min_start_angle_point = min(example_1.new_required_data_angle[0], example_2.new_required_data_angle[0],
                               example_3.new_required_data_angle[0], example_4.new_required_data_angle[0],
                               example_5.new_required_data_angle[0], acc.new_required_data_angle[0])

    a1 = example_1.new_required_data_angle[0] - min_start_angle_point
    a2 = example_2.new_required_data_angle[0] - min_start_angle_point
    a3 = example_3.new_required_data_angle[0] - min_start_angle_point
    a4 = example_4.new_required_data_angle[0] - min_start_angle_point
    a5 = example_5.new_required_data_angle[0] - min_start_angle_point
    a_acc = acc.new_required_data_angle[0] - min_start_angle_point

    for i in range(len(example_1.new_required_data_angle)):
        example_1.new_required_data_angle[i] = example_1.new_required_data_angle[i] - a1

    for i in range(len(example_2.new_required_data_angle)):
        example_2.new_required_data_angle[i] = example_2.new_required_data_angle[i] - a2

    for i in range(len(example_3.new_required_data_angle)):
        example_3.new_required_data_angle[i] = example_3.new_required_data_angle[i] - a3

    for i in range(len(example_4.new_required_data_angle)):
        example_4.new_required_data_angle[i] = example_4.new_required_data_angle[i] - a4

    for i in range(len(example_5.new_required_data_angle)):
        example_5.new_required_data_angle[i] = example_5.new_required_data_angle[i] - a5

    for i in range(len(acc.new_required_data_angle)):
        acc.new_required_data_angle[i] = acc.new_required_data_angle[i] - a_acc

##########################################################################

#

####### rebuild the accelermeter for each machine sensor #################
    golden_sensor_1 = []
    golden_sensor_2 = []
    golden_sensor_3 = []
    golden_sensor_4 = []
    golden_sensor_5 = []

    index_1 = []
    index_2 = []
    index_3 = []
    index_4 = []
    index_5 = []

    for i in range(len(example_1.new_required_data_time)):
        diff_1 = []
        for j in range(len(acc.new_required_data_time)):
            diff_1.append(abs(acc.new_required_data_time[j]-example_1.new_required_data_time[i]))
        index_1.append(diff_1.index(min(diff_1)))
    for i in range(len(example_1.new_required_data_time)):
        golden_sensor_1.append(acc.new_required_data_angle[index_1[i]])
    # print golden_sensor_1

    for i in range(len(example_2.new_required_data_time)):
        diff_2 = []
        for j in range(len(acc.new_required_data_time)):
            diff_2.append(abs(acc.new_required_data_time[j]-example_2.new_required_data_time[i]))
        index_2.append(diff_2.index(min(diff_2)))
    for i in range(len(example_2.new_required_data_time)):
        golden_sensor_2.append(acc.new_required_data_angle[index_2[i]])
    # print golden_sensor_2

    for i in range(len(example_3.new_required_data_time)):
        diff_3 = []
        for j in range(len(acc.new_required_data_time)):
            diff_3.append(abs(acc.new_required_data_time[j]-example_3.new_required_data_time[i]))
        index_3.append(diff_3.index(min(diff_3)))
    for i in range(len(example_3.new_required_data_time)):
        golden_sensor_3.append(acc.new_required_data_angle[index_3[i]])
    # print golden_sensor_3

    for i in range(len(example_4.new_required_data_time)):
        diff_4 = []
        for j in range(len(acc.new_required_data_time)):
            diff_4.append(abs(acc.new_required_data_time[j]-example_4.new_required_data_time[i]))
        index_4.append(diff_4.index(min(diff_4)))
    for i in range(len(example_4.new_required_data_time)):
        golden_sensor_4.append(acc.new_required_data_angle[index_4[i]])
    # print golden_sensor_4

    for i in range(len(example_5.new_required_data_time)):
        diff_5 = []
        for j in range(len(acc.new_required_data_time)):
            diff_5.append(abs(acc.new_required_data_time[j]-example_5.new_required_data_time[i]))
        index_5.append(diff_5.index(min(diff_5)))
    for i in range(len(example_5.new_required_data_time)):
        golden_sensor_5.append(acc.new_required_data_angle[index_5[i]])
    # print golden_sensor_5
    ################################################################

    # plt.plot(example_1.new_required_data_time, example_1.new_required_data_angle, 'r')
    # plt.plot(example_2.new_required_data_time, example_2.new_required_data_angle, 'g')
    # plt.plot(example_3.new_required_data_time, example_3.new_required_data_angle, 'b')
    # plt.plot(example_4.new_required_data_time, example_4.new_required_data_angle, 'y')
    # plt.plot(example_5.new_required_data_time, example_5.new_required_data_angle, 'm')
    # plt.plot(acc.new_required_data_time, acc.new_required_data_angle, 'k')

    plt.scatter(golden_sensor_1, example_1.new_required_data_angle, marker='*', color='r')
    # plt.show()
    # plt.scatter(golden_sensor_2, example_2.new_required_data_angle, marker='*', color='g')
    # plt.show()
    # plt.scatter(golden_sensor_3, example_3.new_required_data_angle, marker='*', color='b')
    # plt.show()
    # plt.scatter(golden_sensor_4, example_4.new_required_data_angle, marker='*', color='y')
    # plt.show()
    # plt.scatter(golden_sensor_5, example_5.new_required_data_angle, marker='*', color='m')
    # plt.show()

###########################################################################################3

    fit_1, y_1 = np.polyfit(golden_sensor_1, example_1.new_required_data_angle, 1)
    fit_1 = float(fit_1)
    y_1 = float(y_1)
    y = []

    for i in golden_sensor_1:
        y.append(fit_1*i + y_1)
    plt.plot(golden_sensor_1,y)

    # fit_2= np.polyfit(golden_sensor_2, example_2.new_required_data_angle, 1)
    # fit_fn_2 = np.poly1d(fit_2)
    # print fit_fn_2

    # fit_3 = np.polyfit(golden_sensor_3, example_3.new_required_data_angle, 1)
    # fit_fn_3 = np.poly1d(fit_3)
    # print fit_fn_3

    # fit_4 = np.polyfit(golden_sensor_4, example_4.new_required_data_angle, 1)
    # fit_fn_4 = np.poly1d(fit_4)
    # print fit_fn_4

    # fit_5 = np.polyfit(golden_sensor_5, example_5.new_required_data_angle, 1)
    # fit_fn_5 = np.poly1d(fit_5)
    # print fit_fn_5

    plt.show()

if __name__ == "__main__":
    main()


