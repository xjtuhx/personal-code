# -*- encoding:gbk -*-

import re

text = u'��4ҳ'
pattern = re.compile(u'��(%d)ҳ')
m = pattern.match(text)
if m:
    print m.group(0)
