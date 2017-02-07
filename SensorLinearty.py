#!/usr/bin/python
import xlrd
import matplotlib.pyplot as plt
import numpy as np


class SensorLinearty():

    def __init__(self):

        self.required_data_time = []
        self.required_data_angle = []
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
        self.required_data_time[:] = self.required_data_time[self.startPoint:]
        # print self.new_required_data_time
        self.required_data_angle[:] = self.required_data_angle[self.startPoint:]
        # print self.new_required_data_angle

    def ajust_angle(self):
        for i in range(len(self.required_data_angle)):
            if (self.required_data_angle[i]-self.required_data_angle[i-1]) > 250:
                self.required_data_angle[i] = self.required_data_angle[i] - 360
            elif (self.required_data_angle[i]-self.required_data_angle[i-1]) < -250:
                self.required_data_angle[i] = self.required_data_angle[i] + 360

    def adjust_stop_time(self):
        counter = 0
        for i in range(len(self.required_data_angle)):
            if abs(self.required_data_angle[i+1] - self.required_data_angle[i]) < 2:
                counter = counter + 1
                if counter > 5:
                    self.endPoint = i
                    break
            else:
                counter = 0
        self.required_data_time[:] = self.required_data_time[:self.endPoint]
        self.required_data_angle[:] = self.required_data_angle[:self.endPoint]

    def shift_start_time(self, min_time):
        shift_time_value = self.required_data_time[0] - min_time
        for i in range(len(self.required_data_time)):
            self.required_data_time[i] = self.required_data_time[i] - shift_time_value

    def shift_start_angle(self, min_angle):
        shift_angle_value = self.required_data_angle[0] - min_angle
        for i in range(len(self.required_data_angle)):
            self.required_data_angle[i] = self.required_data_angle[i] - shift_angle_value

    def make_golden_sensor(self, acc_time, acc_angle, golden_sensor):
        index = []
        for i in range(len(self.required_data_time)):
            diff = []
            for j in range(len(acc_time)):
                diff.append(abs(acc_time[j] - self.required_data_time[i]))
            index.append(diff.index(min(diff)))
        for i in range(len(self.required_data_time)):
            golden_sensor.append(acc_angle[index[i]])

    def make_fit_function(self, golden_sensor, fit, r):
        slop, y = np.polyfit(golden_sensor, self.required_data_angle, 1)
        slop = float(slop)
        y = float(y)
        SStot_all = []
        SSres_all = []

        mean_1 = np.mean(self.required_data_angle)
        # print mean_1

        for i in golden_sensor:
            fit.append(slop * i + y)

        for i in range(len(golden_sensor)):
            SStot_all.append((self.required_data_angle[i] - mean_1) ** 2)
            SSres_all.append((self.required_data_angle[i] - fit[i]) ** 2)
        SStot = sum(SStot_all)
        SSres = sum(SSres_all)
        r = 1 - (SSres / SStot)
        print r

def main():
    mechineSensor_testPathFile = "C:/Users/HPuser/Desktop/sensorLinearty/machine.xlsx"
    accSensor_testPathFile = "C:/Users/HPuser/Desktop/sensorLinearty/acc.xlsx"
    machine_1 = SensorLinearty()
    machine_2 = SensorLinearty()
    machine_3 = SensorLinearty()
    machine_4 = SensorLinearty()
    machine_5 = SensorLinearty()
    acc = SensorLinearty()
    golden_sensor_1 = []
    golden_sensor_2 = []
    golden_sensor_3 = []
    golden_sensor_4 = []
    golden_sensor_5 = []
    fit_1 = []
    fit_2 = []
    fit_3 = []
    fit_4 = []
    fit_5 = []
    r_1 = r_2 = r_3 = r_4 = r_5 = 0

    testWorksheet_1 = machine_1.open_data_file(mechineSensor_testPathFile)
    testWorksheet_2 = machine_2.open_data_file(mechineSensor_testPathFile)
    testWorksheet_3 = machine_3.open_data_file(mechineSensor_testPathFile)
    testWorksheet_4 = machine_4.open_data_file(mechineSensor_testPathFile)
    testWorksheet_5 = machine_5.open_data_file(mechineSensor_testPathFile)
    testWorksheet_acc = acc.open_data_file(accSensor_testPathFile)

    machine_1.get_data(testWorksheet_1, 1, 3)
    machine_1.adjust_start_time()
    machine_1.ajust_angle()
    machine_1.adjust_stop_time()

    machine_2.get_data(testWorksheet_2, 5, 7)
    machine_2.adjust_start_time()
    machine_2.ajust_angle()
    machine_2.adjust_stop_time()

    machine_3.get_data(testWorksheet_3, 9, 11)
    machine_3.adjust_start_time()
    machine_3.ajust_angle()
    machine_3.adjust_stop_time()

    machine_4.get_data(testWorksheet_4, 13, 15)
    machine_4.adjust_start_time()
    machine_4.ajust_angle()
    machine_4.adjust_stop_time()

    machine_5.get_data(testWorksheet_5, 17, 19)
    machine_5.adjust_start_time()
    machine_5.ajust_angle()
    machine_5.adjust_stop_time()

    acc.get_data(testWorksheet_acc,8, 14)
    acc.adjust_start_time()
    acc.ajust_angle()
    acc.adjust_stop_time()
######### ajust the start point of time #################
    min_start_time_point = min(machine_1.required_data_time[0], machine_2.required_data_time[0],
                               machine_3.required_data_time[0], machine_4.required_data_time[0],
                               machine_5.required_data_time[0], acc.required_data_time[0])

    machine_1.shift_start_time(min_start_time_point)
    machine_2.shift_start_time(min_start_time_point)
    machine_3.shift_start_time(min_start_time_point)
    machine_4.shift_start_time(min_start_time_point)
    machine_5.shift_start_time(min_start_time_point)
    acc.shift_start_time(min_start_time_point)

####### adjust the start point in angle #################
    min_start_angle_point = min(machine_1.required_data_angle[0], machine_2.required_data_angle[0],
                               machine_3.required_data_angle[0], machine_4.required_data_angle[0],
                               machine_5.required_data_angle[0], acc.required_data_angle[0])

    machine_1.shift_start_angle(min_start_angle_point)
    machine_2.shift_start_angle(min_start_angle_point)
    machine_3.shift_start_angle(min_start_angle_point)
    machine_4.shift_start_angle(min_start_angle_point)
    machine_5.shift_start_angle(min_start_angle_point)
    acc.shift_start_angle(min_start_angle_point)

######################### make golden sensor for machine 1 ###########################
    machine_1.make_golden_sensor(acc.required_data_time,acc.required_data_angle, golden_sensor_1)
    machine_2.make_golden_sensor(acc.required_data_time, acc.required_data_angle, golden_sensor_2)
    machine_3.make_golden_sensor(acc.required_data_time, acc.required_data_angle, golden_sensor_3)
    machine_4.make_golden_sensor(acc.required_data_time, acc.required_data_angle, golden_sensor_4)
    machine_5.make_golden_sensor(acc.required_data_time, acc.required_data_angle, golden_sensor_5)

######################### make fit function for the plot and calculate the r^2 for machine 1 #################################

    # machine_1.make_fit_function(golden_sensor_1, fit_1, r_1)
    machine_2.make_fit_function(golden_sensor_2, fit_2, r_2)
    # machine_3.make_fit_function(golden_sensor_3, fit_3, r_3)
    # machine_4.make_fit_function(golden_sensor_4, fit_4, r_4)
    # machine_5.make_fit_function(golden_sensor_5, fit_5, r_5)


    ################################################################

    # plt.plot(machine_1.required_data_time, machine_1.required_data_angle, 'r')
    # plt.plot(machine_2.required_data_time, machine_2.required_data_angle, 'g')
    # plt.plot(machine_3.required_data_time, machine_3.required_data_angle, 'b')
    # plt.plot(machine_4.required_data_time, machine_4.required_data_angle, 'y')
    # plt.plot(machine_5.required_data_time, machine_5.required_data_angle, 'm')
    # plt.plot(acc.required_data_time, acc.required_data_angle, 'k')

    # plt.scatter(golden_sensor_1, machine_1.required_data_angle, marker='*', color='r')
    # plt.show()
    plt.scatter(golden_sensor_2, machine_2.required_data_angle, marker='*', color='g')
    # plt.show()
    # plt.scatter(golden_sensor_3, machine_3.required_data_angle, marker='*', color='b')
    # plt.show()
    # plt.scatter(golden_sensor_4, machine_4.required_data_angle, marker='*', color='y')
    # plt.show()
    # plt.scatter(golden_sensor_5, machine_5.required_data_angle, marker='*', color='m')
    # plt.show()

    # plt.plot(golden_sensor_1, fit_1)
    plt.plot(golden_sensor_2, fit_2)
    # plt.plot(golden_sensor_3, fit_3)
    # plt.plot(golden_sensor_4, fit_4)
    # plt.plot(golden_sensor_5, fit_5)

    plt.show()


if __name__ == "__main__":
    main()


