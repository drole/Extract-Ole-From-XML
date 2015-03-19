#SYNTAX: python scriptname.py <xml file>
#OUTPUT: out.bin

import xml.dom.minidom
import binascii
import zlib
import sys

def get_ole(filepath):
    f = open(filepath,'r')
    data = f.read()
    doc = xml.dom.minidom.parseString(data)
    f.close()
    for u in doc.getElementsByTagName('*'):
        if u.firstChild and u.firstChild.nodeValue:
            try:
                data = binascii.a2b_base64(u.firstChild.nodeValue)
            except:
                data = ''
        if data.startswith('ActiveMime'):
            content = zlib.decompress(data[50:])
    return content
    
if __name__ == '__main__':
    if len(sys.argv) > 1:
        print sys.argv[1]
        g = get_ole(sys.argv[1])
        
        fh = file('out.bin','wb')
        fh.write(g)
        fh.close()