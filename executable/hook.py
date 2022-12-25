''' Adds our bin folder to sys.path on launch so we can
    hide our .dll and .pyd files in alternate folders. '''

import sys
import os

CWD = os.path.dirname(sys.argv[0])
sys.path.append(os.path.join(CWD, 'bin'))
