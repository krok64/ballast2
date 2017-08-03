import configparser
import sys

import numpy as np
from mdvlib.util import str_to_arr_rus_float

config = configparser.ConfigParser()
config.read(sys.argv[1], encoding='utf-8')

#координата x
x = str_to_arr_rus_float(config["Coord"]["x"])
#координата y
y = str_to_arr_rus_float(config["Coord"]["y"])
#координаты x для которых надо найти y
x_cut = str_to_arr_rus_float(config["Coord"]["x_cut"])

z = np.polyfit(x, y, 6)
p=np.poly1d(z)
for i in x_cut:
    print(i, "%.3f" % p(i))

