import logging
import datetime
from pytz import timezone


def create_logger(debug, file, stream, full):
    formatter = logging.Formatter('%(name)s - %(funcName)15s - %(lineno)d - %(levelname)s - %(message)s')

    filelog = logging.getLogger('status')
    filelog.setLevel(logging.DEBUG)

    ch = logging.StreamHandler()
    ch.setFormatter(formatter)

    fh = logging.FileHandler(filename='apt.log', mode='w')
    fh.setFormatter(formatter)

    if stream or full:
        ch.setLevel(logging.DEBUG)
    else:
        ch.setLevel(logging.INFO)

    if file or full:
        fh.setLevel(logging.DEBUG)
    else:
        fh.setLevel(logging.DEBUG)

    filelog.addHandler(fh)
    filelog.addHandler(ch)
    return filelog


def init(debug, file, stream, full):
    log = create_logger(debug, file, stream, full)
    fmt = '%Y-%m-%d | %H:%M:%S'

    log.info('==============================================================================')
    log.info('----------------          STARTING-UP THIS BIATCH         --------------------')
    log.info('-------------  BICYCLE GROUP AIR-PERM TEST, POST PROCESSING  -----------------')
    log.info('----------------           {}          --------------------'.format(
        str(datetime.datetime.now(timezone('US/Eastern')).strftime(fmt))
    ))
    log.info('==============================================================================')
    return log
