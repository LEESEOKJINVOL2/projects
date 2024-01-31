# Python SCPI socket functions
# This is not an official Keysight driver.  
# Very limited testing has been done.
# Feel free to modify this
# Version 0.5 


import socket

def SCPI_sock_connect(ipaddress,port=5025):
    """ Opens up a socket connection between an instrument and your PC
        Returns the socket session

        Arguments:
        ipaddress -> ip address of the instrument
        port -> optional -> socket port of the instrument (default 5025)"""

    try:
        session=socket.socket(socket.AF_INET,socket.SOCK_STREAM)
        session.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 0)
        #session.setsockopt(socket.SOL_SOCKET, socket.SO_LINGER, 0)

        session.settimeout(10)
        session.connect((ipaddress, port))
        #session.settimeout(None)

    except socket.timeout:
        print("The socket has not been connecting")
        session = None
    except IOError:
        print ("Failed to connect to the instrument, pleace check your IP address")
        session = None

    return session

def SCPI_sock_send(session,command,error_check=False):
    """Sends a command to an instrument

        Arguments:
        session -> TCPIP socket connection
        command -> text containing an instrument command
        error_check -> optional -> Check for instrument errors (default False)"""

    try:
        bRet = True
        strCmd = command + '\n'
        bCmd = to_bytes(strCmd)
        session.sendall(bCmd)

        if error_check is True:
            err = get_error(session, command)
            if err:
                bRet = False
        return bRet

    except socket.timeout:
        print("Send error:")
        return False

    except Exception as err:
        print(err)
        return False

        
def SCPI_sock_query(session,command,error_check=False):
    """Sends a query to an instrument
        Returns the query response
        
        Arguments:
        session -> TCPIP socket connection
        command -> text containing an instrument command
        error_check -> optional -> Check for instrument errors (default False)"""
    
    session.settimeout(60)
    try:
        strCmd = command + '\n'
        bCmd = to_bytes(strCmd)
        session.sendall(bCmd)
        response = getDataFromSocket(session)
        if error_check is True:
            err = get_error(session, command)
            if err:
                response = None
        return response
        
    except socket.timeout:
        print ("Query error:")
        get_error(session, command)
        return None

    except Exception as err:
        return None

def SCPI_sock_close(session):
    """Closes the socket connection

        Argument:
        session -> TCPIP socket connection"""
    
    session.close()

def getDataFromSocket(session):
    """Reads from a socket until a newline is read
        Returns the data read

        Argument:
        session -> TCPIP socket"""
    
    dat = ""
    while 1:
        message = session.recv(4096)
        last=len(message)
        if message[last-1] == 0x0A:
            dat = message[:-1]
        else:
            dat = message

        retData = to_str(dat)

        return(retData)

def get_error(session, command):
    """Checks an instrument for errors and print them out
        Returns True if any errors are encountered

        Arguments:
        session -> TCPIP socket connection
        command -> text containing an instrument command"""
        
    has_err=False
    resp = SCPI_sock_query(session,"SYST:ERR?")
    
    if int(resp[:2]) != 0:
        print ("Your command: " + command + " has errors:")
        print (resp)
        has_err = True
    while int(resp[:2]) != 0:
        resp = SCPI_sock_query(session,"SYST:ERR?")
        if int(resp[:2]) != 0:
            print(resp)

    return has_err

def to_bytes(bytes_or_str):
    if isinstance(bytes_or_str, str):
        value = bytes_or_str.encode() # uses 'utf-8' for encoding
    else:
        value = bytes_or_str
    return value # Instance of bytes


def to_str(bytes_or_str):
    if isinstance(bytes_or_str, bytes):
        value = bytes_or_str.decode() # uses 'utf-8' for encoding
    else:
        value = bytes_or_str
    return value # Instance of str

def SCPI_check_connection(session):
    if session is None:
        return 1

    rsp_data = SCPI_sock_query(session, '*IDN?')
    if rsp_data is None:
        return 2
    else:
        print(to_str(rsp_data))
        return 0