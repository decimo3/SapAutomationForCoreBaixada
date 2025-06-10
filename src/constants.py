''' Module to hold constant values '''
#region imports
import os
import sys
import dotenv
#endregion

SEPARADOR = ';'
SHORT_TIME_WAIT = 5
LONG_TIME_WAIT = 10
DESIRE_INSTANCES = 5

if getattr(sys, 'frozen', False):
  BASE_FOLDER = os.path.dirname(sys.executable)
else:
  BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))

dotenv.load_dotenv(os.path.join(BASE_FOLDER, 'sap.conf'))

LOCKFILE = os.path.join(BASE_FOLDER, 'sap.lock')

NOTUSE = str(os.environ.get('NOTUSE')).split(',')

EXCLUDE = str(os.environ.get('EXCLUDE')).split(',')
