# -*- encoding:utf-8 -*-
from HTMLParser import HTMLParser
import re
import urllib2
import datetime
import PyRSS2Gen

URL = "http://anonomous.yculblog.com/"

class PostGetter(HTMLParser):
    def __init__(self, href):
        HTMLParser.__init__(self)
        self.title = ""
        self.author = ""
        self.timestamp = ""
        self.content = ""
        self.href = href
        self.titlefound = False
        self.authorfound = False
        self.timestampfound = False
        self.contentfound = False

    def handle_starttag(self, tag, attrs):
        if self.contentfound == True:
            self.content += self.get_starttag_text()
        if tag == 'a':
            for k,v in attrs:
                if k == 'class' and v == 'post_title':
                    self.titlefound = True
        elif tag == 'span':
            for k,v in attrs:
                if k == 'class' and v == 'post_user':
                    self.authorfound = True
                elif k == 'class' and v == 'post_time':
                    self.timestampfound = True
        elif tag == 'div':
            for k,v in attrs:
                if k == 'class' and v == 'post_content':
                    self.contentfound = True

    def handle_endtag(self, tag):
        if tag == 'a' and self.titlefound == True:
            self.titlefound = False
        elif tag == 'span':
            if self.authorfound == True:
                self.authorfound = False
            if self.timestampfound == True:
                self.timestampfound = False
        elif tag == 'div' and self.contentfound == True:
            self.contentfound = False
        if self.contentfound == True:
            self.content += ("</"+tag+">")

    def handle_data(self, data):
        if self.titlefound == True:
            self.title += data
        if self.authorfound == True:
            self.author += data
        if self.timestampfound == True:
            self.timestamp += data
        if self.contentfound == True:
            self.content += data

    def getRssItem(self):
        return PyRSS2Gen.RSSItem(
                title = self.title,
                link = self.href,
                description = self.content,
                guid = PyRSS2Gen.Guid(self.href),
                pubDate = datetime.datetime.strptime(
                    self.timestamp.strip('@ \t\n')
                    , "%Y-%m-%d %H:%M"))

class PostFinder(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.postfound = False
        self.rssitems =[]
        self.pattern = re.compile('共(\d+)页', re.UNICODE)
        self.pages = 0
    
    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            for k,v in attrs:
                if k == 'href':
                    if re.match('post\.\d+\.html', v):
                        self.postfound = True
                        f = urllib2.urlopen(URL + v)
                        content = unicode(f.read(), 'UTF-8')
                        f.close()
                        postgetter = PostGetter(URL + v)
                        postgetter.feed(content)
                        postgetter.close()
                        self.rssitems.append(postgetter.getRssItem())
                        print '<a href="',v,'">'

    def handle_endtag(self, tag):
        if tag == 'a' and self.postfound == True:
            self.postfound = False
            print '</a>'

    def handle_data(self, data):
        m = self.pattern.match(data)
        if m:
            print "match found"
            self.pages = int(m.group(1))
        if self.postfound == True:
            print data

    def getRssItems(self):
        return self.rssitems

    def getPages(self):
        return self.pages

if __name__=='__main__':
    parser = PostFinder()
    for i in range(1, 5):
        f = urllib2.urlopen(URL + ("archive.p%d.html" % i))
        content = unicode(f.read(), "utf-8")
        f.close()
        parser.feed(content)
        parser.close()
    rss = PyRSS2Gen.RSS2(
            title = "Faith and Calm",
            link = URL,
            description = "",
            lastBuildDate = datetime.datetime.now(),
            items = parser.getRssItems())
    rss.write_xml(open("anonomous.xml", "w"), encoding="utf-8")
