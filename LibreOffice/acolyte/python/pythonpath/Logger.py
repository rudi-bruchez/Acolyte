# -*- coding: utf-8 -*-
import os
import uno
import unohelper
from com.sun.star.document import XEventListener
from com.sun.star.task import XJobExecutor
from openai import OpenAI
from pathlib import Path
import logging

#====== Logging
log_all_disabled = False
my_log_level = 10 #'CRITICAL':50, 'ERROR':40, 'WARNING':30, 'INFO':20, 'DEBUG':10, 'NOTSET':0


class Logger:
    def __init__(self, doc):
        """
        Initialize the Logger class.

        Args:
            doc: The LibreOffice document.
        """
        this_doc_url = doc.URL
        this_doc_sys_path = uno.fileUrlToSystemPath(this_doc_url)
        this_doc_parent_path = Path(this_doc_sys_path).parent
        base_name = Path(this_doc_sys_path).name
        only_name = str(base_name).rsplit('.ods')[0]
        log_file_path = Path(this_doc_parent_path, f'''{only_name}.log''')

        logger = logging.getLogger(__name__)
        logger.setLevel(my_log_level)
        logger.disabled = log_all_disabled

        fh = logging.FileHandler(str(log_file_path), mode='w')
        fh.setLevel(my_log_level)
        formatter = logging.Formatter(fmt='%(asctime)s:%(levelname)s:%(filename)s:%(lineno)d:%(message)s', datefmt='%d-%m-%Y:%H:%M:%S')
        fh.setFormatter(formatter)
        logger.addHandler(fh)

    def log(self, message):
        """
        Log a message.

        Args:
            message (str): The message to log.
        """
        logger = logging.getLogger(__name__)
        logger.log(my_log_level, message)