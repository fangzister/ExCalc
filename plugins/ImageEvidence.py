# coding: utf-8

import os
import exifread
import codecs
from hashlib import md5
import sys

reload(sys)
sys.setdefaultencoding('gb2312')


def getExif(filename):
    FIELD = 'EXIF DateTimeOriginal'
    fd = open(filename, 'rb')
    tags = exifread.process_file(fd)
    fd.close()
    t = str(tags[FIELD])
    return t.replace(':', '-', 2)


def main():
    if len(sys.argv) == 2:
        dd = sys.argv[1]
    else:
        dd = 'D:\\ExCalc\\plugins\\'

    d = open(dd + 'filelist.txt', 'r')
    f = codecs.open(dd + 'list.txt', 'w', encoding='gb2312')
    i = 0
    f.write(u'序号\t名称\t文件类型\t拍摄时间\t物理大小(字节)\tMD5值\n')

    for filename in d.read().splitlines():
        if os.path.isfile(filename) and filename.lower().endswith('.jpg'):
            i += 1
            nm = os.path.basename(filename)
            tm = getExif(filename)
            ss = format(os.path.getsize(filename), ',')
            m = md5()
            p = open(filename, 'rb')
            m.update(p.read())
            p.close()
            m5 = m.hexdigest()
            f.write('%d\t%s\t%s\t%s\t%s\t%s\n' % (i, nm, u'jpg图片', tm, ss, m5))
            # print filename,tm
    f.close()


main()

