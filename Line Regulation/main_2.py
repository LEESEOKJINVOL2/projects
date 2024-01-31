import os
import csv
from ctrl_keysight_n6705b import *
from smu24xx import SMU24xx
import subprocess

def generate_ps_voltage(start, end, step):
    ps_voltage = []
    current_value = start
    while current_value >= end:
        ps_voltage.append(round(current_value, 1))
        current_value -= step
    return ps_voltage

PS_VOLTAGE = generate_ps_voltage(4.8, 2.9, 0.1)
SMU_CURRENT = ['0', '-30e-3', '-60e-3', '-90e-3', '-120e-3', '-150e-3']

reset_script = "python /Users/cm/Documents/pydiag/laguna_reset.py"#reset cript path

class Main():
    def __init__(self):
        self.n6705b = ctrl_keysight_n6705b()#To declare keysight library
        self.n6705b.connection('192.168.100.50', '5025')#To Connect Power Supply
        self.n6705b.set_port(True, '1')#To set port for PS
        self.n6705b.set_port(True, '2')#To set port 2 for charge sense
        self.smu = SMU24xx('192.168.100.10')#To Connect SMU

    def func_call_script(self, call_script_cmd):#run reset script("python /Users/cm/Documents/pydiag/laguna_reset.py")
        ret, err = self.subprocess_open(call_script_cmd)
        return ret

    def subprocess_open(self, command):
        popen = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        (stdoutdata, stderrdata) = popen.communicate()
        return stdoutdata, stderrdata

    def run(self):
        # turn on EVB when PGM run
        self.n6705b.set_voltage(4.8, '1')
        self.r_si, self.r_sv = self.smu.test_run2(str(SMU_CURRENT[0]))
        self.n6705b.set_voltage(4.0, '2')
        time.sleep(1)
        self.n6705b.set_voltage(0, '2')
        print("PGM run")

        for j, cur in enumerate(SMU_CURRENT):  # to input smu current step
            self.r_si, self.r_sv = self.smu.test_run2(str(SMU_CURRENT[j]))  # To sink current and measure voltage from SMU
            time.sleep(1)
            for i, vol in enumerate(PS_VOLTAGE):# Get Power Supply voltage value
                self.n6705b.set_voltage(vol, '1')#send volatage and ch# to PS
                time.sleep(1)
                self.func_call_script(reset_script)#to reset laguna
                time.sleep(1)
                self.r_pv = self.n6705b.read_volt(1)#To read voltage from PS
                time.sleep(0.2)
                self.r_pi = self.n6705b.read_curr(1)#To read current from PS
                time.sleep(0.2)
                eff = (abs((self.r_sv * self.r_si)) / abs((self.r_pv * self.r_pi))) * 100#to calculate efficiency
                print(f'Vdd Main: {vol}V, Sink Curr: {float(cur)*1000}mA, SMU Voltage: {round(self.r_sv,3)}V,Effi: {round(eff,2)}%')
                self.save_data(self.r_pv, self.r_pi, self.r_sv, self.r_si, eff, float(cur))#to save data
                time.sleep(1)
            
            self.n6705b.set_voltage(4.8, '1')  # port 1: 3.8V(default voltage for EVB Vdd main)
            self.r_si, self.r_sv = self.smu.test_run2(str(SMU_CURRENT[0])) # SMU current 0
            self.n6705b.set_voltage(4.0, '2') # port 2: 3.95V
            time.sleep(1)
            print("CHG SNS: reset")
            self.n6705b.set_voltage(0, '2')  # port 2: 0V
            print(f"{cur}: cycle done")

    def save_data(self, r_pv, r_pi, r_sv, r_si, eff, loadi):# to output measured data
        header_list = ['PS_VOLT', 'PS_CURR', 'SMU_VOLT', 'SMU_CURR', 'EFFICIENCY', 'LOAD_CURR']
        save_list = [round(r_pv, 1), r_pi, r_sv, r_si, eff, loadi] # PS voltage round off(ex: 4.8, 3.8 etc in csv file)
        file_name = 'result/data.csv'#output file name
        csvfile_is = os.path.isfile(file_name)
        if csvfile_is is False:
            with open(file_name, 'w') as File:
                csv_writer = csv.writer(File)
                csv_writer.writerow(header_list)
                csv_writer.writerow(save_list)
        else:
            with open(file_name, 'a') as File:
                csv_writer = csv.writer(File)
                csv_writer.writerow(save_list)

if __name__ == '__main__':
    a = Main()
    a.run()
    print('PyCharm')