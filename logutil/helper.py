from logging import getLogger
from logging import StreamHandler
from logging import Formatter
from logging import NOTSET
from logging import DEBUG

def setup_logger(level=NOTSET):
    logger = getLogger()
    sh = StreamHandler()
    formatter = Formatter('%(asctime)s [%(levelname)s] %(name)s > %(message)s')
    sh.setFormatter(formatter)
    logger.addHandler(sh)
    logger.setLevel(level)
    sh.setLevel(level)
    return logger

    