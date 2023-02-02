# coding:utf-8
import nfc
import binascii
import errno
import contextlib
import ctypes
import threading
import os
import sys
sys.path.insert(1, os.path.split(sys.path[0])[0])

service_code = 0x090f
stu_num = ''

def connected(tag):
    # タグのIDなどを出力する
    #print (tag)
    global stu_num

    if isinstance(tag, nfc.tag.tt3.Type3Tag):
        try:
          # 内容を16進数で出力する
          #print('  ' + '\n  '.join(tag.dump()))
          stu_num = tag.dump()[3].split('|')[1][:10]
        except Exception as e:
          print ("error: %s" % e)
          stu_num = 'error'
          return stu_num
    else:
      print ("error: tag isn't Type3Tag")
      stu_num = 'error'
    return stu_num


class TimeoutException(IOError):
    errno = errno.EINTR

@contextlib.contextmanager
def time_limit_with_thread(timeout_secs):
    thread_id = ctypes.c_ulong(threading.get_ident())
    def raise_exception():
        modified_thread_state_nums = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, ctypes.py_object(TimeoutException))
        if modified_thread_state_nums == 0:
            raise ValueError('Invalid thread id. thread_id:{}'.format(thread_id))
        elif modified_thread_state_nums > 1:
            # Normally do not go through this path, but clear unthrown Exceptions just in case
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, None)
            raise SystemError('PyThreadState_SetAsyncExc failure.')

    timer = threading.Timer(timeout_secs, raise_exception)
    timer.setDaemon(True)
    timer.start()
    try:
        yield
    finally:
        timer.cancel()
        timer.join()

def access_card():
    clf = nfc.ContactlessFrontend('usb')
    #print('touch card:')
    try:
        tag = clf.connect(rdwr={'on-connect': lambda tag: False})
    finally:
        clf.close()
    idm = binascii.hexlify(tag.idm)
    return idm

def get_idm():
    timeout_secs = 0.7
    with time_limit_with_thread(timeout_secs):
        try:
            idm = access_card()
        except:
            idm = ''
    return idm

def get_stunum():
    timeout_secs = 1
    idm = str(get_idm())
    idm_list = ["b'01101800091c3f01'", "b'01160400f919bb1e'"]
    if idm in idm_list:
        if idm == "b'01101800091c3f01'":
            idm = 'shimazaki'
        elif idm == "b'01160400f919bb1e'":
            idm = '1181201166'
        print('idm', idm)
        return idm
    with time_limit_with_thread(timeout_secs):
        global stu_num
        stu_num = ''
        try:
            clf = nfc.ContactlessFrontend('usb')
            clf.connect(rdwr={'on-connect': connected})
        except:
            stu_num = ''
    return stu_num

if __name__ == '__main__':
    #print(get_idm())
    #stunum = get_stunum()
    print(get_idm())
