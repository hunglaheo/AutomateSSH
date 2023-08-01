import telnetlib
import time
import json
import re
import sys
tn = telnetlib.Telnet("10.153.15.35",23)
response = tn.read_until(b'login:')
print(response.decode())
tn.write(b'rfprodsdc\n')
response = tn.read_until(b'word:')
print(response.decode())
tn.write(b'rfprodsdc\n')
response = tn.read_until(b'word:')
print(response.decode())