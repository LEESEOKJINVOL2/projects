import socketscpi as scpi
import time

SCPI_CMD = [
    '*RST',                                 # Reset
    'SENS:FUNC "CURR"',                     # Set to measure current.
    'SENS:CURR:RANG:AUTO ON',               # Set to measure with Auto range enabled.
    'SENS:CURR:RSEN ON',                    # Set to use 4-wire sense mode.
    'SOUR:FUNC VOLT',                       # Set to source voltage.
    'SOUR:VOLT:RANG:AUTO ON',               # Set to voltage range.
    'SOUR:VOLT:ILIM 1',                     # Set the current limit to 1A.
    'SOUR:SWE:VOLT:LIN -2, 2, 100, 0.01',   # Set to sweep voltage from -2V to 2V in 1000steps with a 0.1s
    ':INIT',                                # Initiate the sweep.
]
MODEL = [
    'MODEL SIM',
    'MODEL 2450',
    'MODEL 2460',
]
AB_SPEC = {
    'MODEL 2460': (105, 7,)
}

"""
SOURCE_RANGING =[
Best Fixed,
Auto,
Fixed,
]
"""

class SMU24xx:
    def __init__(self, ip=''):
        try:
            self.scpi = scpi.SocketInstrument(ip,timeout=600) #time out 10sec
            print(self.scpi.instId)
            self.model = self.scpi.instId.split(',')[1]
            if self.model in MODEL:
                print(self.model)
            else:
                raise Exception('There is no model.')

            # cmd = 'TRAC:DATA? 1, {}, "defbuffer1", SOUR, READ'.format(20)
            # print(self.scpi.query(cmd))
            # instr = ':SOUR:VOLT:READ:BACK?'
            # print(self.scpi.query(instr))

            self.cmd = []
            self.enable = True
            self.dly = .0
            self.swp_v = {}
            self.set_sweep_voltage(-2, 2, 100, 0.1)
            self.four_wire_mode = True
            self.sensing_auto = True
            self.sensing_range = 0.001
            self.force_limit = 0.001
            # self.err = ''

        except Exception as e:
            print(e)
            raise Exception('[SMU2450] Please check connection and IP Address({})'.format(ip))

    def info(self):
        return self.scpi.instId

    def reset(self):
        self.scpi.write('*RST')

    def run(self):
        pass
        # self.reset()
        # time.sleep(0.1)  # Delay
        # self.set_iv_mode()
        # time.sleep(0.1)  # Delay
        # self.start()


    def start(self):
        self.scpi.write(':INIT')  # Initiate the sweep. START

    def wait_cmd(self):
        # 이전 실행한 cmd가 완료되기 전까지 다음 커맨드를 실행하지 않게 하는 명령어
        self.scpi.write('*WAI')

    def set_wire_mode(self, wire):
        if wire == '2-Wire':
            self.four_wire_mode = False
        else:
            self.four_wire_mode = True


    def set_sensing_mode(self,mode, sens_range, force_limit):
        if mode == 'Manual':
            self.sensing_auto = False
            self.sensing_range = sens_range
        else:
            self.sensing_auto = True
            self.sensing_range = 0

        self.force_limit = force_limit

    def set_sweep_voltage(self, start, stop, step, dly):
        self.swp_v['start_v'] = start
        self.swp_v['stp_v'] = stop
        self.swp_v['step'] = step
        self.swp_v['dly'] = dly
        # s = 'SOUR:SWE:VOLT:LIN {}, {}, {}, {}'.format(
        #     self.swp_v['start_v'], self.swp_v['stp_v'], self.swp_v['step'], self.swp_v['dly'])
        # print(s)

    def get_sweep_voltage(self):
        s = 'SOUR:SWE:VOLT:LIN {}, {}, {}, {}, 1, FiX, OFF'.format(
            self.swp_v['start_v'], self.swp_v['stp_v'], self.swp_v['step'], self.swp_v['dly'])
        return s

    def set_iv_mode(self):
        dly=0.01
        self.scpi.write('SENS:FUNC "CURR"')  # Set to measure current.
        time.sleep(dly)
        if self.sensing_auto:
            self.scpi.write('SENS:CURR:RANG:AUTO ON')  # Set to measure with Auto range enabled.
        else:
            self.scpi.write('SENS:CURR:RANG {}'.format(self.sensing_range))  # Set to measure with Auto range enabled.
        time.sleep(dly)
        if self.four_wire_mode:
            self.scpi.write('SENS:CURR:RSEN ON')  # Set to use 4-wire sense mode.
        else:
            self.scpi.write('SENS:CURR:RSEN OFF')  # Set to use 4-wire sense mode.
        time.sleep(dly)
        self.scpi.write('SOUR:FUNC VOLT')  # Set to source voltage.
        time.sleep(dly)
        self.scpi.write('SOUR:VOLT:RANG:AUTO ON')  # Set to voltage range.
        time.sleep(dly)
        self.scpi.write(':OUTP:SMOD zero')  # zero ouput state
        time.sleep(dly)
        self.scpi.write('SOUR:VOLT:ILIM {}'.format(self.force_limit))  # Set the current limit to 0.0002A.
        time.sleep(dly)
        self.scpi.write(self.get_sweep_voltage())
        # time.sleep(0.05)
        # self.scpi.write(':INIT')  # Initiate the sweep. START

    def set_zero_output_state(self, delay=1):
        dly = 0.01
        self.reset()
        time.sleep(dly)
        self.scpi.write(':OUTP:VOLT:SMOD ZERO')  # zero ouput state
        time.sleep(dly)
        self.scpi.write('SENS:FUNC "CURR"')  # Set to measure current.
        time.sleep(dly)
        self.scpi.write('SENS:CURR:RANG:AUTO ON')  # Set to measure with Auto range enabled.
        time.sleep(dly)
        self.scpi.write('SOUR:FUNC VOLT')  # Set to source voltage.
        time.sleep(dly)
        self.scpi.write('SOUR:VOLT:RANG:AUTO ON')  # Set to voltage range.
        time.sleep(dly)
        self.scpi.write('SOUR:VOLT -0.1')  # Set to voltage range.
        time.sleep(dly)
        self.scpi.write('SOUR:VOLT:ILIM 0.2')  # Set the current limit to 0.001A.
        time.sleep(dly)
        self.scpi.write(':OUTP ON')
        time.sleep(delay)
        self.scpi.write(':OUTP OFF')

    def test_run(self, start, stop, step, end):
        dly = 0.01
        self.reset()
        self.scpi.write('SENS:FUNC "VOLT"')
        time.sleep(dly)
        self.scpi.write('SENS:VOLT:RANG:AUTO ON')
        time.sleep(dly)

        self.scpi.write('SOUR:FUNC CURR')
        time.sleep(dly)
        self.scpi.write('SOUR:CURR:RANG:AUTO ON')
        time.sleep(dly)
        self.scpi.write('SOUR:CURR:VLIM 5')
        time.sleep(dly)

        self.scpi.write(':OUTP:SMOD zero')
        time.sleep(dly)
        self.scpi.write('SOUR:SWE:CURR:LIN {}, {}, {}, {}, 1, FiX, OFF'.format(start, stop, step, end))
        time.sleep(dly)

        self.scpi.write(':INIT')
        time.sleep(dly)
        self.scpi.write('*WAI')
        time.sleep(dly)
        i, v = self.get_iv_result(step)
        return i, v

    def test_run2(self, curr):#To control SMU
        dly = 0.2
        self.reset()
        self.scpi.write('SENS:FUNC "VOLT"')
        time.sleep(dly)
        self.scpi.write('SENS:VOLT:RANG:AUTO ON')
        time.sleep(dly)

        self.scpi.write('SOUR:FUNC CURR')
        time.sleep(dly)
        self.scpi.write('SOUR:CURR:RANG:AUTO ON')
        time.sleep(dly)
        self.scpi.write('SOUR:CURR:VLIM 5')
        time.sleep(dly)
        self.scpi.write('SOUR:CURR {}'.format(curr))
        time.sleep(dly)

#Sourcing action
        time.sleep(dly)
        self.smu_on()
        time.sleep(dly)
        i = float(self.scpi.query('MEAS:CURR?'))
        time.sleep(dly)
        v = float(self.scpi.query('MEAS:VOLT?'))
        time.sleep(dly)
#Read result

        return i, v#return measured current and voltage from SMU

    def smu_on(self):
        self.scpi.write(':OUTP ON')
    def smu_off(self):
        self.scpi.write(':OUTP OFF')
    def is_done(self, sec=10):
        retv = ''
        start = time.time()
        while time.time() - start < sec:
            try:
                retv = self.scpi.query('*OPC?')
                print('*OPC?')
            except :
                retv = '0'
                time.sleep(0.05)
            if retv == '1':
                break
        return retv


    def get_iv_result(self, step):
        #cmd = 'TRAC:DATA? 1, 100, "defbuffer1", TSTamp, SOUR, READ'
        cmd = 'TRAC:DATA? 1, {}, "defbuffer1", SOUR, READ'.format(step)
        print(cmd)

        t = self.scpi.query(cmd)
        if t == '1':
            t = self.scpi.query(cmd)

        a = []
        v = []
        for i, s in enumerate(t.split(','), start=1):
            if i % 2 == 1:
                v.append(s) # insert current measure
            else:
                a.append(s) # insert voltage

        return v, a



if __name__ == "__main__":

    ip = '192.168.0.10'
    SMU = SMU24xx(ip)

    #SMU.set_wire_mode('2-Wire')
    #SMU.set_sensing_mode('Auto', )
    v, a = SMU.run()


    SMU.set_zero_output_state()

    #v, a = SMU.get_iv_result()

