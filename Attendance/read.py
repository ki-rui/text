from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

import logging
import os
import time

def log():
    logger = logging.getLogger()
    logger.setLevel(level=logging.DEBUG)
    time_line = time.strftime('%Y%m%d%H%M', time.localtime(time.time()))
    log_path=os.path.join(os.getcwd(),'log')
    if not os.path.exists(log_path):
        os.mkdir(log_path)
    logfile = log_path + '/'+time_line + '.txt'
    handler = logging.FileHandler(logfile,mode='w') # 输出到log文件的handler
    # handler.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger
if __name__=='__main__':
    l=log()
    l.debug('This is a debug message.')
    l.info('This is an info message.')
    l.warning('This is a warning message.')
    l.error('This is an error message.')
    l.critical('This is a critical message.')

