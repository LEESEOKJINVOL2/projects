import os, sys
from SCPI_socket import *
import time
import math
# import logging
import numpy

class ctrl_keysight_n6705b:
    def __init__(self):
        self.session = 0


        return

    def check_connection(self, session):
        if session is None:
            # logging.info('Socket was not connected!')
            # self.logger.info('Socket was not connected!')
            return False
        self.rsp_data = SCPI_sock_query(session, '*IDN?')
        if self.rsp_data is None:
            # print('No response!')
            # self.logger.info('No response!')
            return False
        else:
            pass
            # print(to_str(self.rsp_data))
            # self.logger.info(to_str(self.rsp_data))
        return True


    def connection(self, ip, port):
        self.session = SCPI_sock_connect(ip, int(port))
        if (self.check_connection(self.session) == False):
            return False
        return True

    def set_voltage(self, volt, ch):
        cmd = 'VOLT {0},(@{1});'.format(volt, ch)
        SCPI_sock_send(self.session, cmd)
        # logging.info('set_voltage {0}'.format(cmd))
        # self.logger.info('set_voltage {0}'.format(cmd))

    def set_current(self, curr, ch):
        cmd = 'CURR {0},(@{1});'.format(curr, ch)
        SCPI_sock_send(self.session, cmd)
        # self.logger.info('set_current {0}'.format(cmd))

    def set_port(self, on, ch):
        if(on == True):
            cmd = 'OUTPut ON,(@{0});'.format(ch)
        else:
            cmd = 'OUTPut OFF,(@{0});'.format(ch)
        SCPI_sock_send(self.session, cmd)
        # logging.info('set_Port {0}'.format(cmd))
        # self.logger.info('set_Port {0}'.format(cmd))


    def run(self,volt,ch):
        self.set_voltage(volt,ch)
        self.set_port(True,ch)


    def read_curr(self,ch):
        list_current = []

        for i in range(10):
            cmd = 'MEAS:CURR? (@{0})'.format(ch)
            ret = SCPI_sock_query(self.session, cmd)
            ret = float("{:.8f}".format(float(ret)))
            list_current.append(ret)

        avg = round(numpy.average(list_current), 5)
        # self.logger.info('read_curr {0}'.format(avg))
        return avg

    def read_volt(self,ch):
        list_voltage = []

        for i in range(10):
            cmd = 'MEAS:VOLT? (@{0})'.format(ch)
            ret = SCPI_sock_query(self.session, cmd)
            ret = float("{:.8f}".format(float(ret)))
            list_voltage.append(ret)

        avg = round(numpy.average(list_voltage), 5)
        # self.logger.info('read_curr {0}'.format(avg))
        return avg




if __name__ == "__main__":
    print('main function: n6705b')
    n6705b = ctrl_keysight_n6705b()
    if(n6705b.connection('10.86.44.211', 5025) == True):
        n6705b.set_voltage(2,1)
        n6705b.read_curr(1)
        n6705b.set_port(False,1)