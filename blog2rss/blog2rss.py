from xml.sax import saxutils
from xml.sax import handler
from xml.sax import make_parser
from xml.sax.handler import feature_namespaces
import urllib2
import re

class FindPost(handler.ContentHandler, handler.ErrorHandler):
    def __init__(self, title):
        self.search_title = title

    def startElement(self, name, attrs):
        if name != 'a': return
        if re.match('post\.\d+\.html', attrs['href']):
            print '<a href="',attrs['href'],'"/>'

    def error(self, exception):
        import sys
        sys.stderr.write("%s\n" % exception)

    def fatalError(self, exception):
        import sys
        sys.stderr.write("%s\n" % exception)

if __name__=='__main__':
    parser = make_parser()
    parser.setFeature(feature_namespaces, 0)

    dh = FindPost('a')
    parser.setContentHandler(dh)

    f = urllib2.urlopen('http://anonomous.yculblog.com/archive.html')
    parser.parse(f)
    f.close()
