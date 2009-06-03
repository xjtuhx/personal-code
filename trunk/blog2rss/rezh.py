# -*- encoding:gbk -*-

import re

text = u'共4页'
pattern = re.compile(u'共(%d)页')
m = pattern.match(text)
if m:
    print m.group(0)
