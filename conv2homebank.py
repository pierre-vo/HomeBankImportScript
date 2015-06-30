#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Author: Pierre
# Purpose: process files exported from bank website to inport in HomeBank
###############################################################################
#
#
###############################################################################
"""@package docstring

"""

# Rq: CSV output has more info

# Todo: fix unicode issues with french and german characters

from __future__ import absolute_import
from __future__ import unicode_literals  # all strings in this file are unicode

import sys
import os
import logging
import re
import time
import csv
import json
import tempfile
import argparse
from os.path import isfile, join

from codecs import open

logger = logging.getLogger(__name__)

# Payment modes handled by HomeBank (csv codes)
PAYMODES = ["None",  # 0
            "Credit Card",  # 1
            "Check",  # 2
            "Cash",  # 3
            "Transfer",  # 4
            "Internal Transfer",  # 5
            "Debit Card",  # 6
            "Standing Order",  # 7
            "Electronic Payment",  # 8
            "Deposit",  # 9
            "FI Fees"]  # 10

# for automatic detection of files in "In" directory and automatic processing
filetypes = {'INGDiba_csv': r'Umsatzanzeige_[0-9]{10}_[0-9]{8}.*.csv',
             'Boursorama_qif': r'[0-9]{11}_Q[0-9]{8}.*.qif',
             'Boursorama_quick2000': r'[0-9]{11}_R[0-9]{8}.*.qif',
             'Linxo_csv': r'op\xe9rations.csv'
             }


def dispdic(in_d):
    print json.dumps(in_d, sort_keys=True,
                     indent=4, separators=(',', ': '))


def listExtFromDir(ext, in_dir):
    file_l = [f for f in (in_dir) if isfile(join(in_dir, f))]
    file_l = [f for f in file_l if (f.find(ext) != -1)]
    return file_l


class boursorama_qif_file:
    """ Handles the QIF files from boursorama output
    Data representation of internal dict is messy... """
    def __init__(self, in_file):
        self.in_file = in_file
        self.op_d = {}
        self.headerline = None

        self.open_qif()
        self.process_op()

    def open_qif(self):
        logger.info('* opening {}'.format(self.in_file))
        with open(self.in_file, 'rb') as fid:
            self.headerline = fid.readline().replace('Ccard', 'CCard')
            fbuff = fid.read()

        # Replace weird chars
#         fbuff = re.sub(u"\u007B", u"\xe9", fbuff)
#         fbuff = re.sub(u"\u00C2\u00A0", "", fbuff)

        op_l = fbuff.split('^')
        self.op_d = self.read_op_l(op_l)

    def read_op_l(self, op_l):
        logger.debug('=> {} operations in file.'.format(len(op_l)))
        out_d = {}
        for op in op_l:
            op_d = self.read_op(op.rstrip('\n').lstrip('\n'))
            if len(op_d) > 1:
                ts = time.strptime(op_d['Date'], '%m/%d/%Y')
                ts = int(time.mktime(ts))
                while ts in out_d:
                    ''' If another record with the same timestamp exists,
                    add 1 sec since the timestamp will be used as
                    dictionnary key'''
                    ts += 1
                out_d[ts] = op_d
            else:
                logger.error('! Empty record ?')
                logger.error('! => {}'.format(op_d))
        return out_d

    def read_op(self, op_l):
#         logger.debug('\t* {}'.format(op_l.replace('\n', '\t|')))
        op_d = {}
        l = op_l.split('\n')

        for item in l:
            if len(item) > 0:
                if item[0] == 'D':
                    op_d['Date'] = item[1:].replace("'", "/20")
                elif item[0] == 'T':
                    op_d['Montant'] = float(item[1:].replace(',', ''))
                elif item[0] == 'P':
                    op_d['Descr'] = item[1:]
                else:
                    logger.error('! Categorie non prise en'
                                 ' charge: {}'.format(item))

#         logger.debug('\t  => {}'.format(op_d))
        return op_d

    def process_op(self):
        for key in sorted(self.op_d.keys()):
            l1 = re.match(r'^([A-Z \.]*)([0-9]{6})\s*([A-Z0-9]{2})(.*)',
                          self.op_d[key]['Descr'])
            self.op_d[key]['Parse'] = {}
            self.op_d[key]['Parse']['date'] = self.op_d[key]['Date']\
                                                        .replace('/', '')
            self.op_d[key]['Parse']['lieu'] = ''
            if l1 != None:
                self.op_d[key]['Parse']['type'] = l1.group(1).rstrip(' ')\
                                                             .lstrip(' ')
                self.op_d[key]['Parse']['date'] = l1.group(2).rstrip(' ')\
                                                             .lstrip(' ')
                self.op_d[key]['Parse']['lieu'] = l1.group(3).rstrip(' ')\
                                                             .lstrip(' ')
                self.op_d[key]['Parse']['Descr'] = l1.group(4).rstrip(' ')\
                                                              .lstrip(' ')
            elif self.op_d[key]['Descr'].find('VIR SEPA') != -1:
                self.op_d[key]['Parse']['type'] = 'VIR SEPA'
                self.op_d[key]['Parse']['Descr'] = self.op_d[key]['Descr']\
                                                    .replace('VIR SEPA ', '')
            elif self.op_d[key]['Descr'].find('PRLV SEPA') != -1:
                self.op_d[key]['Parse']['type'] = 'PRLV SEPA'
                self.op_d[key]['Parse']['Descr'] = self.op_d[key]['Descr']\
                                                    .replace('PRLV SEPA ', '')
            elif self.op_d[key]['Descr'].find('CHQ') != -1:
                self.op_d[key]['Parse']['type'] = 'CHQ.'
                self.op_d[key]['Parse']['Descr'] = self.op_d[key]['Descr']
                del1 = self.op_d[key]['Descr'].find('N.')
                self.op_d[key]['Parse']['NrCheque'] = \
                                            self.op_d[key]['Descr'][del1 + 2:]
            elif self.op_d[key]['Descr'].find('RETRAIT DAB') != -1:
                self.op_d[key]['Parse']['type'] = 'RETRAIT'
                self.op_d[key]['Parse']['Descr'] = self.op_d[key]['Descr']
            elif self.op_d[key]['Descr'].find('VIR') != -1:
                self.op_d[key]['Parse']['type'] = 'VIR'
                self.op_d[key]['Parse']['Descr'] = self.op_d[key]['Descr']
            elif self.op_d[key]['Descr'].find('PRLV') != -1:
                self.op_d[key]['Parse']['type'] = 'PRLV'
                self.op_d[key]['Parse']['Descr'] = self.op_d[key]['Descr']
            elif self.op_d[key]['Descr'].find('Relev') != -1:
                self.op_d[key]['Parse']['type'] = 'Releve Carte'
                del1 = self.op_d[key]['Descr'].find('Carte')
                self.op_d[key]['Parse']['Descr'] = 'Relev\xe9 diff\xe9r\xe9 '\
                                               + self.op_d[key]['Descr'][del1:]
            else:
                logger.error('! Ligne ignor\xe9e:')
                logger.error('! {}'.format(self.op_d[key]['Descr']))
                self.op_d[key]['Parse']['type'] = '?'
                self.op_d[key]['Parse']['Descr'] = self.op_d[key]['Descr']

    def dic2HBdic(self):
        out_d = {}
        # HB fields: date, paymode, info, payee, memo, amount, category, tags

        for item in self.op_d:
            sub_d = {'date': self.op_d[item]['Date'],
                     'paymode': PAYMODES.index("None"),
                     'info': None,
                     'payee': self.op_d[item]['Parse']['Descr'],
                     'memo': self.op_d[item]['Descr'],
                     'amount': self.op_d[item]['Montant'],
                     'category': None,
                     'tags': None}

            if self.op_d[item]['Parse']['type'] == 'PAIEMENT CARTE':
                sub_d['paymode'] = PAYMODES.index("Credit Card")
                sub_d['info'] = 'CB'
            elif self.op_d[item]['Parse']['type'].find('CHQ') != -1:
                sub_d['paymode'] = PAYMODES.index("Check")
                sub_d['info'] = self.op_d[item]['Parse']['NrCheque']
            elif self.op_d[item]['Parse']['type'].find('VIR') != -1:
                sub_d['paymode'] = PAYMODES.index("Transfer")
            elif self.op_d[item]['Parse']['type'].find('PRLV') != -1:
                sub_d['paymode'] = PAYMODES.index("Standing Order")
            elif self.op_d[item]['Parse']['type'].find('RETRAIT') != -1:
                sub_d['paymode'] = None
                sub_d['info'] = None

            out_d[item] = sub_d
        return out_d


class ING_DiBa_csv_file:
    def __init__(self, in_file):
        self.in_file = in_file
        self.op_d = {}
        self.headerline = None

        self.open_csv()

    def open_csv(self):
        logger.info('* opening {}'.format(self.in_file))
        # remove header from csv file
        with open(self.in_file, 'rb', encoding="cp1252") as fid:
            fbuff = fid.read()

        fbuff = fbuff[fbuff.find("Buchung"):]
        # Replace weird chars
#         fbuff = re.sub(u"\u007B", u"\xe9", fbuff)
#         fbuff = re.sub(u"\u00C2\u00A0", "", fbuff)
#         fbuff = re.sub(u"\xE4", "a", fbuff)
#         fbuff = re.sub(u"\xDC", "U", fbuff)

        # save temp csv file to process
        fid = tempfile.NamedTemporaryFile(delete=False)
        fid.write(fbuff.encode('utf8'))
        fid.close()

        with open(fid.name, 'rb') as csvfile:
            csvr = csv.reader(csvfile, delimiter=';'.encode('utf8'),
                              quotechar='"'.encode('utf8'))

            for row in csvr:
                if self.headerline is None:
                    self.headerline = row
                    logger.debug('=> {} elements in header.'\
                                 .format(len(self.headerline)))
                    # clean header
                    for idx in range(len(self.headerline)):
                        self.headerline[idx] = self.headerline[idx].decode('utf-8').rstrip('"')
                        logger.info(self.headerline[idx])
                else:
                    self.read_op_l(row)
        os.unlink(fid.name)
        logger.debug('=> {} operations processed.'.format(len(self.op_d)))

    def read_op_l(self, op_l):
        sub_d = {}
        for item in self.headerline:
            if item in ["Buchung", "Valuta"]:
                sub_d[item] = time.strptime(op_l[self.headerline.index(item)]\
                                            .replace('"', '')\
                                            .rstrip(' '), '%d.%m.%Y')
                sub_d[item] = time.mktime(sub_d[item])
            elif item in ["Betrag", "Saldo"]:
                sub_d[item] = float(op_l[self.headerline.index(item)]\
                                    .replace('"', '').rstrip(' ')\
                                    .replace('.', '').replace(',', '.'))
            else:
                sub_d[item] = op_l[self.headerline.index(item)].decode('utf-8').rstrip(' ')

        key = sub_d["Buchung"]
        # make sure we have no duplicated keys in the dict
        while key in self.op_d:
            key += 1
        self.op_d[key] = sub_d

    def dic2HBdic(self):
        out_d = {}
        # HB fields: date, paymode, info, payee, memo, amount, category, tags
        for item in self.op_d:
            sub_d = {'date': time.strftime('%m/%d/%Y',
                                   time.localtime(self.op_d[item]['Buchung'])),
                     'paymode': PAYMODES.index("None"),
                     'info': None,
                     'payee': self.op_d[item]['Auftraggeber/Empfänger'],
                     'memo': self.op_d[item]['Verwendungszweck'],
                     'amount': self.op_d[item]['Betrag'],
                     'category': None,
                     'tags': None}

            conv = {'Lastschrifteinzug': PAYMODES.index("Credit Card"),
                    'Uberweisung': PAYMODES.index("Transfer"),
                    'Gutschrift': PAYMODES.index("Transfer"),  # 5 = problem?
                    'Gutschrift aus Dauerauftrag': PAYMODES.index("Transfer"),
                    'Dauerauftrag/Terminueberweisung': PAYMODES.\
                                                            index("Transfer"),
                    }
            if self.op_d[item]['Buchungstext'] in conv:
                sub_d['paymode'] = conv[self.op_d[item]['Buchungstext']]

            out_d[item] = sub_d
        return out_d


class linox_csv_file:
    def __init__(self, in_file):
        self.in_file = in_file
        self.op_d = {}
        self.headerline = None

        self.open_csv()

    def open_csv(self):
        logger.info('* opening {}'.format(self.in_file))
        with open(self.in_file, 'rb', encoding='utf-16') as f:
            for row in csv.reader(f, delimiter='\t'.encode('ascii')):
                if self.headerline == None:
                    self.headerline = row
                    logger.debug('=> {} elements in header.'\
                                 .format(len(self.headerline)))
                else:
                    if len(row) == len(self.headerline):
                        self.read_op_l(row)
                    else:
                        logger.error('! Empty record?')
                        logger.error('! {}'.format(row))

        logger.info('=> {} operations processed.'.format(len(self.op_d)))

    def read_op_l(self, op_l):
        sub_d = {}
        for item in self.headerline:
            if item == "Date":
                sub_d[item] = time.strptime(op_l[self.headerline.index(item)],
                                            '%d/%m/%Y')
                sub_d[item] = time.mktime(sub_d[item])
            elif item == "Montant":
                sub_d[item] = float(op_l[self.headerline.index(item)]\
                                    .replace('"', '').rstrip(' ')\
                                    .replace('.', '').replace(',', '.'))
            else:
                sub_d[item] = op_l[self.headerline.index(item)].rstrip(' ')

        key = sub_d["Date"]
        # make sure we have no duplicated keys in the dict
        while key in self.op_d:
            key += 1
        self.op_d[key] = sub_d

    def dic2HBdic(self):
        out_d = {}
        # HB fields: date, paymode, info, payee, memo, amount, category, tags
        # unicode issues with labels, taking them from self.headline
        # ['Date', 'Libell\xc3\xa9', 'Cat\xc3\xa9gorie', 'Montant', 'Notes', 'N\xc2\xb0 de ch\xc3\xa8que', 'Labels']
        for item in self.op_d:
#             dispdic(self.op_d[item])
            sub_d = {'date': time.strftime('%m/%d/%Y',
                                   time.gmtime(self.op_d[item]['Date'])),
                     'paymode': PAYMODES.index("None"),
                     'info': None,
                     'payee': self.op_d[item][self.headerline[1]],
                     'memo': self.op_d[item][self.headerline[1]],
                     'amount': self.op_d[item]['Montant'],
                     'category': None,
                     'tags': None}

#             conv = {'Lastschrifteinzug': PAYMODES.index("Credit Card"),
#                     'Uberweisung': PAYMODES.index("Transfer"),
#                     'Gutschrift': PAYMODES.index("Internal Transfer"),
#                     'Gutschrift aus Dauerauftrag': PAYMODES.index("Transfer"),
#                     'Dauerauftrag/Terminueberweisung': PAYMODES.\
#                                                             index("Transfer"),
#                     }
#             if self.op_d[item]['Buchungstext'] in conv:
#                 sub_d['paymode'] = conv[self.op_d[item]['Buchungstext']]

            out_d[item] = sub_d
        return out_d


class HomeBankDataWriter:
    def __init__(self, in_dic, head=None):
        ''' Import dict containing data '''
        logger.debug('* Import data dic with {} records.'.format(len(in_dic)))
        self.op_d = in_dic
        self.headerline = head

    def export_qif(self, out_file):
        ''' Export to a QIF format handled by Homebank '''
        cnt = 0
        if self.headerline == None:
            logger.warning('! No QIF header defined, using default.')
            self.headerline = '!Type:Bank\n'
        with open(out_file, 'wb') as fid:
            fid.write(self.headerline)
            for key in sorted(self.op_d.keys()):
                self.write_op(fid, self.op_d[key])
                fid.write('^\n')
                cnt += 1
        logger.info('=> Exported {} entries to {}.'.format(cnt, out_file))

    def write_op(self, fid, op_d):
        ''' Write QIF operation '''
        fid.write('D{}\n'.format(op_d['date']).encode('utf8'))
        fid.write('T{}\n'.format(op_d['amount']).encode('utf8'))
        fid.write('P{}\n'.format(op_d['payee']).encode('utf8'))
        fid.write('M{}\n'.format(op_d['memo']).encode('utf8'))

    def export_csv(self, out_file):
        cnt = 0
        try:
            with open(out_file, 'wb') as fid:
                head_l = ['date', 'paymode', 'info', 'payee', 'wording',
                          'amount', 'category', 'tags']

                dic_l = ['date', 'paymode', 'info', 'payee', 'memo',
                         'amount', 'category', 'tags']

                # .encode("codec") should be applied to str before writing
                # need to determing code expected by homebank
                fid.write(';'.join(head_l))
                fid.write('\n')

                for key in sorted(self.op_d.keys()):
                    l = []
                    for item in dic_l:
                        t = self.op_d[key][item]
                        if self.op_d[key][item] == None:
                            l.append('')
                        elif (isinstance(t, int) or isinstance(t, float)):
                            l.append(str(self.op_d[key][item]))
                        elif isinstance(t, unicode):
                            l.append(self.op_d[key][item])
                        elif isinstance(t, str):
                            l.append(self.op_d[key][item].decode('utf-8'))
                        else:
                            logger.error('!!! Type of element appended not defined => will cause trouble')
                            logger.error('!!! Type: {} ({})'.format(type(t), t))
                            l.append(self.op_d[key][item])
                    logger.info(l)
                    try:
                        fid.write(';'.join(l))
                        fid.write('\n')
                        cnt += 1
                    except BaseException, e:
                        logger.error('!!! {} - {}'.format(e, l))
        except IOError:
            logger.error('! Cannot open the file,'
                         ' could it be opened in Excel ?')
            logger.error('! {}'.format(out_file))
        logger.info('=> Exported {} entries to {}.'.format(cnt, out_file))


def getTypeFromFileName(in_file):
    for typ in filetypes:
        if re.match(filetypes[typ], in_file):
            return typ
    return None


def main_no_args():
    logger.debug('Start')
    in_dir = 'In'
    out_dir = 'Out'

    file_l = [f for f in os.listdir(in_dir) if isfile(join(in_dir, f))]

    for f in file_l:
        typ = getTypeFromFileName(f)
        in_file = join(in_dir, f)
        out_file_csv = join(out_dir, f.replace('.qif', '.csv'))
        out_file_qif = join(out_dir, f.replace('.csv', '.qif'))

        logger.info('')
        logger.info('** In: {}'.format(in_file))
        if typ == 'INGDiba_csv':
            data_d = ING_DiBa_csv_file(in_file)
            HB = HomeBankDataWriter(data_d.dic2HBdic())
            HB.export_qif(out_file_qif)
            HB.export_csv(out_file_csv)
        elif typ == 'Boursorama_qif':
            data_d = boursorama_qif_file(in_file)
            HB = HomeBankDataWriter(data_d.dic2HBdic(), data_d.headerline)
            HB.export_qif(out_file_qif)
            HB.export_csv(out_file_csv)
        elif typ == 'Linxo_csv':
            logger.warning('! Implementation not complete')
            data_d = linox_csv_file(in_file)
            HB = HomeBankDataWriter(data_d.dic2HBdic())
            HB.export_qif(out_file_qif)
            HB.export_csv(out_file_csv)
        else:
            logger.error('! Type not determined for {}'.format(f))
            logger.error('! Skipping.')


def main():
    logging.basicConfig(level=logging.DEBUG,
                        format='[%(levelname)-5s] %(lineno)s - %(message)s',
                        datefmt='%M:%S')

    p = argparse.ArgumentParser()
    p.add_argument('-i', '--input', help="input file")
    types = filetypes.keys()
    p.add_argument('-t', '--type', choices=types, help="Type of the file")
    # sys.getfilesystemencoding()

    args = p.parse_args()
    print args

    if args.input == None:
        logger.info('* No arguments, attempting to automatically process "In"')
        main_no_args()
    else:
        logger.info('* In: {}'.format(args.input))
        if not isfile(args.input):
            logger.error('! {} is not a valid file.'.format(args.input))
            raise ValueError
        if args.type == None:
            logger.info('* No type defined, trying to determine')
            typ = getTypeFromFileName(args.input)
        else:
            typ = args.type

        in_file = args.input
        out_file_csv = os.path.abspath(in_file.replace('.qif', 'conv.csv'))
        out_file_qif = os.path.abspath(in_file.replace('.csv', 'conv.qif'))

        if typ == 'INGDiba_csv':
            data_d = ING_DiBa_csv_file(in_file)
            HB = HomeBankDataWriter(data_d.dic2HBdic())
            HB.export_qif(out_file_qif)
            HB.export_csv(out_file_csv)
        elif typ == 'Boursorama_qif':
            data_d = boursorama_qif_file(in_file)
            HB = HomeBankDataWriter(data_d.dic2HBdic(), data_d.headerline)
            HB.export_qif(out_file_qif)
            HB.export_csv(out_file_csv)
        elif typ == 'Linxo_csv':
            logger.warning('! Implementation not complete')
            data_d = linox_csv_file(in_file)
            HB = HomeBankDataWriter(data_d.dic2HBdic())
            HB.export_qif(out_file_qif)
            HB.export_csv(out_file_csv)
        else:
            logger.error('! Type not determined for {}'.format(in_file))
            logger.error('! Skipping.')


if __name__ == "__main__":
    main()
