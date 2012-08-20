import struct
from collections import namedtuple
import csv, types, codecs, cStringIO

# Example usage:
# from pySTLparse import stlFile
# a_file = stlFile('Affixes.stl')
# a_file.writecsv('output.csv')

_MPQheader_len = 0x10

# For CSV output
class UnicodeWriter:
    def __init__(self, f, dialect=csv.excel, encoding="utf-8", **kwds):
        self.queue = cStringIO.StringIO()
        self.writer = csv.writer(self.queue, dialect=dialect, **kwds)
        self.stream = f
        self.encoder = codecs.getincrementalencoder(encoding)()
    def writerow(self, row):
        temp = []
        for s in row:
            if (type(s) == types.IntType):
                temp.append(str(s))
            else:
                temp.append(s)
        self.writer.writerow([s.encode("utf-8") for s in temp])
        data = self.queue.getvalue()
        data = data.decode("utf-8")
        data = self.encoder.encode(data)
        self.stream.write(data)
        self.queue.truncate(0)
    def writerows(self, rows):
        for row in rows:
            self.writerow(row)

StlString = namedtuple('StlString', ['string1', 'string2', 'string3', 'string4'])

class stlFile:
    def _lookup_str(self, a_stlentry):
        # Translates string offsets into actual string values
        str1 = ''
        if (a_stlentry.string1size > 1):
            str1 = self._stl[_MPQheader_len + a_stlentry.string1offset:_MPQheader_len +
                a_stlentry.string1offset + a_stlentry.string1size - 1]
        str2 = ''
        if (a_stlentry.string2size > 1):
            str2 = self._stl[_MPQheader_len + a_stlentry.string2offset:_MPQheader_len +
                a_stlentry.string2offset + a_stlentry.string2size - 1]
        str3 = ''
        if (a_stlentry.string3size > 1):
            str3 = self._stl[_MPQheader_len + a_stlentry.string3offset:_MPQheader_len +
                a_stlentry.string3offset + a_stlentry.string3size - 1]
        str4 = ''
        if (a_stlentry.string4size > 1):
            str4 = self._stl[_MPQheader_len + a_stlentry.string4offset:_MPQheader_len +
                a_stlentry.string4offset + a_stlentry.string4size - 1]
        return StlString(str1,str2,str3,str4)
    def __init__(self, filename):
        f = open(filename, "rb")
        stl = f.read()
        f.close()
        self._stl = stl
        MPQheader_offset = 0x00
        MPQheader = namedtuple('MPQheader', ['mpqMagicNumber', 'fileTypeId', 'unk_0', 'unk_1'])
        header    = MPQheader(*struct.unpack('<4L', stl[MPQheader_offset:MPQheader_offset+_MPQheader_len]))
        StlHeader_offset = 0x10
        StlHeader_len    = 0x28
        StlHeader = namedtuple('StlHeader', ['stlFileId', 'unk1_0', 'unk1_1', 'unk1_2', 'unk1_3', 'unk1_4',
            'headerSize', 'entriesSize', 'unk2_0', 'unk2_1'])
        stlh = StlHeader(*struct.unpack('<10L', stl[StlHeader_offset:StlHeader_offset+StlHeader_len]))
        StlEntry_len = 0x50
        numEntries   = stlh.entriesSize / StlEntry_len
        StlEntry = namedtuple('StlEntry', ['unk1_0', 'unk1_1', 'string1offset', 'string1size',
                                           'unk2_0', 'unk2_1', 'string2offset', 'string2size',
                                           'unk3_0', 'unk3_1', 'string3offset', 'string3size',
                                           'unk4_0', 'unk4_1', 'string4offset', 'string4size',
                                           'unk5', 'unk6_0', 'unk6_1', 'unk6_2'])
        all_strings = []
        for i in xrange(numEntries):
            start = _MPQheader_len + StlHeader_len + (StlEntry_len * i)
            a_str = StlEntry(*struct.unpack('<20L', stl[start:start+StlEntry_len]))
            all_strings.append(a_str)
        self.strings = []
        for x in all_strings:
            self.strings.append(self._lookup_str(x))
    def writecsv(self, filename='output.csv'):
        outfile = open(filename, 'wb')
        csv.register_dialect('exceltab', delimiter='\t')
        writer = UnicodeWriter(outfile, dialect='exceltab')
        writer.writerow(['string1', 'string2', 'string3', 'string4'])
        for a_string in self.strings:
            writer.writerow(list(a_string))
        outfile.close()
