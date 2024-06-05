import os
import sys

def get_executable_path():
    return os.path.dirname(os.path.abspath(sys.argv[0]))
