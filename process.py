import os
import glob
from Main import *

i=1
path = 'data/input'
for filename in glob.glob(os.path.join(path, '*.xlsx')):
       print(filename)
       main(filename,i)
       i+=1
