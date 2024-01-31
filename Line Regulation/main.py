import os
import csv
#import matplotlib.pyplot as plt
from ctrl_keysight_n6705b import *
from smu24xx import SMU24xx
import subprocess

# PS_VOLTAGE = [4.8,4.3,3.8,3.3,2.9]#input Vdd main array  #[4.8,4.5,4.2,3.8,3.5,3.2,2.9]
PS_VOLTAGE = [2.9]
#PS_VOLTAGE = [2.9]
# SMU_CURRENT = ['0', '-10e-3', '-20e-3', '-30e-3', '-40e-3', '-50e-3'] #line regulation, effi sink current array ['0','-10e-3','-20e-3','-30e-3','-40e-3','-50e-3']
SMU_CURRENT = ['0', '-30e-3', '-60e-3', '-90e-3', '-120e-3', '-150e-3']

reset_script = "python /Users/cm/Documents/pydiag/laguna_reset.py"#reset cript path

class Main():
    def __init__(self):
        self.n6705b = ctrl_keysight_n6705b()#To declare keysight library
        self.n6705b.connection('192.168.100.50', '5025')#To Connect Power Supply
        self.n6705b.set_port(True, '1')#To set port for PS
        self.smu = SMU24xx('192.168.100.10')#To Connect SMU

    def func_call_script(self, call_script_cmd):#run reset script("python /Users/cm/Documents/pydiag/laguna_reset.py")
        ret, err = self.subprocess_open(call_script_cmd)
        return ret

    def subprocess_open(self, command):
        print("print command::" + command)
        popen = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        (stdoutdata, stderrdata) = popen.communicate()
        print(stdoutdata)
        print(stderrdata)
        return stdoutdata, stderrdata

    def run(self):
        for i, vol in enumerate(PS_VOLTAGE):# Get Power Supply voltage value
            self.n6705b.set_voltage(vol, '1')#send volatage and ch# to PS
            time.sleep(1)
            self.func_call_script(reset_script)#to reset laguna
            # time.sleep(1)
            for j, cur in enumerate(SMU_CURRENT):#to input smu current step
                self.r_si, self.r_sv = self.smu.test_run2(str(SMU_CURRENT[j]))#To sink current and measure voltage from SMU
                time.sleep(1)
                self.r_pv = self.n6705b.read_volt(1)#To read voltage from PS
                time.sleep(0.2)
                self.r_pi = self.n6705b.read_curr(1)#To read current from PS
                time.sleep(0.2)
                eff = (abs((self.r_sv * self.r_si)) / abs((self.r_pv * self.r_pi))) * 100#to calculate efficiency
                print(f'Vdd Main: {vol}V, Sink Curr: {float(cur)*1000}mA, Buck Voltage: {round(self.r_sv,3)}V,Effi: {round(eff,2)}%')
                self.save_data(self.r_pv, self.r_pi, self.r_sv, self.r_si, eff, float(cur))#to save data
                self.smu.smu_off()#turn off SMU
                print('#######VOLTAGE {} , CURRENT {} Measument DONE##########'.format(vol, cur))
                time.sleep(1)

        self.n6705b.set_port(False, '1')

    def save_data(self, r_pv, r_pi, r_sv, r_si, eff, loadi):# to output measured data
        header_list = ['PS_VOLT', 'PS_CURR', 'SMU_VOLT', 'SMU_CURR', 'EFFICIENCY', 'LOAD_CURR']
        save_list = [r_pv, r_pi, r_sv, r_si, eff, loadi]
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

def plotLineRegulation_r2():
    data = pd.read_csv('data.csv', low_memory=False, encoding='cp949', sep=",")#데이터 Reading
    data['LOAD_CURR'] = abs(data['SMU_CURR']*1000)
    p1 = data[data['LOAD_CURR'] == 10]
    p2 = data[data['LOAD_CURR'] == 50]
    p3 = data[data['LOAD_CURR'] == 100]
    p4 = data[data['LOAD_CURR'] == 200]
    p5 = data[data['LOAD_CURR'] == 250]
    p6 = data[data['LOAD_CURR'] == 300]
    p1.rename(columns={'LOAD_CURR':'10mA'},inplace=True)
    p2.rename(columns={'LOAD_CURR': '50mA'},inplace=True)
    p3.rename(columns={'LOAD_CURR': '100mA'},inplace=True)
    p4.rename(columns={'LOAD_CURR': '200mA'},inplace=True)
    p5.rename(columns={'LOAD_CURR': '250mA'},inplace=True)
    p6.rename(columns={'LOAD_CURR': '300mA'},inplace=True)

    ax = plt.gca()
    p1.plot(kind='line',x='PS_VOLT',y='10mA',marker = 'o',ax=ax)#2.8V 선 생성
    p2.plot(kind='line',x='PS_VOLT',y='50mA',marker = 's',ax=ax)#2.8V 선 생성
    p3.plot(kind='line',x='PS_VOLT',y='100mA',marker = '*',ax=ax)#2.8V 선 생성
    p4.plot(kind='line',x='PS_VOLT',y='200mA',marker = 'x',ax=ax)#2.8V 선 생성
    p5.plot(kind='line',x='PS_VOLT',y='250mA',marker = 'D',ax=ax)#2.8V 선 생성
    p6.plot(kind='line',x='PS_VOLT',y='300mA',marker = '|',ax=ax)#2.8V 선 생성
    if round(data.iloc[0][2], 1) == 1.5:
        plt.title('Line Regulation Buck4')

        plt.grid(True)# grid 생성

        plt.xlabel('Vddmain [V]')# x축 라벨
        plt.ylabel('Vbuck4 [V]')# y축 라벨

        plt.xlim(2.6, 5.0)  # y축 고정
        plt.ylim(-1, 2)  # y축 고정

        plt.xticks(np.arange(2.8, 5, 0.2))  # x축 간격 설정
        plt.yticks(np.arange(-1.2, 2.2, 0.2))  # y축 간격 설정

        plt.savefig('Line Regulation_Buck4.png',dpi=200)
        plt.show()

    if round(data.iloc[0][2], 1) == 1.8:
        plt.title('Line Regulation Buck5')

        plt.grid(True)  # grid 생성

        plt.xlabel('Vddmain [V]')  # x축 라벨
        plt.ylabel('Vbuck5 [V]')  # y축 라벨

        plt.xlim(2.6, 5.0)  # y축 고정
        plt.ylim(-1, 2)  # y축 고정

        plt.xticks(np.arange(2.8, 5, 0.2))  # x축 간격 설정
        plt.yticks(np.arange(-1.2, 2.2, 0.2))  # y축 간격 설정

        plt.savefig('Line Regulation_Buck5.png', dpi=200)
        plt.show()

    if round(data.iloc[0][2], 1) == 1.2:
        plt.title('Line Regulation Buck3')

        plt.grid(True)  # grid 생성

        plt.xlabel('Vddmain [V]')  # x축 라벨
        plt.ylabel('Vbuck3 [V]')  # y축 라벨

        plt.xlim(2.6, 5.0)  # y축 고정
        plt.ylim(-1, 2)  # y축 고정

        plt.xticks(np.arange(2.8, 5, 0.2))  # x축 간격 설정
        plt.yticks(np.arange(-1.2, 2.2, 0.2))  # y축 간격 설정

        plt.savefig('Line Regulation_Buck3.png', dpi=200)
        plt.show()

    if round(data.iloc[0][2], 1) == 0.7:
        plt.title('Line Regulation Buck1')

        plt.grid(True)  # grid 생성

        plt.xlabel('Vddmain [V]')  # x축 라벨
        plt.ylabel('Vbuck1 [V]')  # y축 라벨

        plt.xlim(2.6, 5.0)  # y축 고정
        plt.ylim(-1, 2)  # y축 고정

        plt.xticks(np.arange(2.8, 5, 0.2))  # x축 간격 설정
        plt.yticks(np.arange(-1.2, 2.2, 0.2))  # y축 간격 설정

        plt.savefig('Line Regulation_Buck1.png', dpi=200)
        plt.show()

    if round(data.iloc[0][2], 1) == 0.6:
        plt.title('Line Regulation Buck0')

        plt.grid(True)  # grid 생성

        plt.xlabel('Vddmain [V]')  # x축 라벨
        plt.ylabel('Vbuck0 [V]')  # y축 라벨

        plt.xlim(2.6, 5.0)  # y축 고정
        plt.ylim(-1, 2)  # y축 고정

        plt.xticks(np.arange(2.8, 5, 0.2))  # x축 간격 설정
        plt.yticks(np.arange(-1.2, 2.2, 0.2))  # y축 간격 설정

        plt.savefig('Line Regulation_Buck0.png', dpi=200)
        plt.show()


if __name__ == '__main__':
    a = Main()
    a.run()
    print('PyCharm')
    #plotLineRegulation_r2()
    #plotLineRegulation_r2()