#!/usr/bin/python
#
#
# resources: see http://wiki.services.openoffice.org/wiki/Python
#

from xmlrpclib import ServerProxy
from random import randint
import nfqueue
from dpkt import ip, tcp
from socket import AF_INET, AF_INET6, inet_ntoa
import sys
import optparse
import os

IDX_IDX = 0
IDX_FORBIDDEN = 1
IDX_ALLOWED = 2
IDX_VERDICT = 3

def get_list_of_tcp_ports():
    ret = dict()
    y = 1
    while True:
        cell_port = oo_instance.getCell(session, book, sheet, 0, y)
        if not cell_port:
            break
        cell_forbidden = oo_instance.getCell(session, book, sheet, 1, y)
        cell_allowed = oo_instance.getCell(session, book, sheet, 2, y)
        cell_verdict = oo_instance.getCell(session, book, sheet, 3, y)
        ret[int(cell_port)] = (y, cell_forbidden, cell_allowed, cell_verdict)
        y += 1
    return ret

def update_stats_for_port(dport, pkt, is_allowed):
    # get index
    val = oo_tcp_ports[dport]
    idx = val[IDX_IDX]
    if is_allowed:
        oo_index = 2
    else:
        oo_index = 1
    # get cell value
    cell_value = oo_instance.getCell(session, book, sheet, oo_index, idx)
    if not cell_value:
        cell_value = 0
    print "previous cell value: ", cell_value
    # increment, then update
    cell_value += 1
    oo_instance.setCell(session, book, sheet, oo_index, idx, cell_value)

def get_list_of_words():
    ret = dict()
    y = 1
    while True:
        cell_word = oo_instance.getCell(session, book, sheet + 1, 0, y)
        if not cell_word:
            break
        cell_translation = oo_instance.getCell(session, book, sheet + 1, 1, y)
        cell_translated = oo_instance.getCell(session, book, sheet + 1, 2, y)
        ret[cell_word] = (y, cell_translation, cell_translated)
        y += 1
    return ret

def update_stats_for_word(word):
    # get index
    val = oo_words[word]
    idx = val[IDX_IDX]
    oo_index = 2
    # get cell value
    cell_value = oo_instance.getCell(session, book, sheet + 1, oo_index, idx)
    if not cell_value:
        cell_value = 0
    print "previous cell value: ", cell_value
    # increment, then update
    cell_value += 1
    oo_instance.setCell(session, book, sheet + 1, oo_index, idx, cell_value)

def cb(i,payload):
        global pkt_counter, oo_tcp_ports, oo_words, options
        # every n packets, re-read list of ports
        if pkt_counter % 50 == 0:
            print "re-reading list of TCP ports"
            oo_tcp_ports = get_list_of_tcp_ports()
            print "re-reading list of words"
            oo_words = get_list_of_words()
        pkt_counter += 1
        #print "payload len ", payload.get_length()
        data = payload.get_data()
        pkt = ip.IP(data)
        decision = nfqueue.NF_DROP
        #print "proto:", pkt.p
        #print "source: %s" % inet_ntoa(pkt.src)
        #print "dest: %s" % inet_ntoa(pkt.dst)
        if pkt.p == ip.IP_PROTO_TCP:
                #print "  sport: %s" % pkt.tcp.sport
                #print "  dport: %s" % pkt.tcp.dport
                dport = pkt.tcp.dport
                if dport in oo_tcp_ports:
                    t = oo_tcp_ports[dport]
                    if t[IDX_VERDICT] == 1:
                        print "Port %d allowed" % dport
                        update_stats_for_port(dport, pkt, 1)
                        decision = nfqueue.NF_ACCEPT
                # don't check for packet return on SYN
                if pkt.tcp.flags & tcp.TH_SYN and not pkt.tcp.flags & tcp.TH_ACK:
                    payload.set_verdict(decision)
                    sys.stdout.flush()
                    return 1
                # crappy reverse check
                sport = pkt.tcp.sport
                if sport in oo_tcp_ports and dport >= 1024:
                    t = oo_tcp_ports[sport]
                    if t[IDX_VERDICT] == 1:
                        print "S Port %d allowed" % sport
                        #update_stats_for_port(dport, pkt, 1)
                        decision = nfqueue.NF_ACCEPT
                    else:
                        print "Port %d dropped" % dport
                        #update_stats_for_port(dport, pkt, 0)
                        payload.set_verdict(nfqueue.NF_DROP)
                        sys.stdout.flush()
                        return 1
                if options.do_substitution:
                    # accepted data packet need to be modified
                    if pkt.tcp.flags & tcp.TH_PUSH:
                        pkt2 = pkt
                        for word in oo_words.keys():
                            if str(pkt.tcp.data).find(word) != -1:
                                # Do substitution
                                print "Found %s word" % word
                                print pkt2.tcp.data
                                old_len = len(pkt2.tcp.data)
                                pkt2.tcp.data = str(pkt2.tcp.data).replace(word,oo_words[word][1])
                                pkt2.len = pkt2.len - old_len + len(pkt2.tcp.data)
                                pkt2.tcp.sum = 0
                                pkt2.sum = 0
                                update_stats_for_word(word)
                        payload.set_verdict_modified(nfqueue.NF_ACCEPT,str(pkt2),len(pkt2))
                    elif decision == nfqueue.NF_DROP:
                        payload.set_verdict(nfqueue.NF_DROP)
        payload.set_verdict(nfqueue.NF_ACCEPT)

        sys.stdout.flush()
        return 1


parser = optparse.OptionParser(usage='oowall.py [-s] ')
parser.add_option("-c", "--config-file", dest="oofile", default=os.environ.get('PWD') + '/oowall.ods',
             help="location of the oocalc file")
parser.add_option("-u", "--uri", dest="uri", default='http://localhost:8000',
             help="URI of the pyuno server")
parser.add_option("-s", "--with-substitution", dest="do_substitution",
             action="store_true", default=False,
             help="activate on-the-fly substitution")


try:
    options, args = parser.parse_args()
except IndexError:
    parser.print_help()
    sys.exit(2)

q = nfqueue.queue()
q.set_callback(cb)
q.fast_open(0, AF_INET)

print "Using %s calc file and connecting to %s" % (options.oofile, options.uri)

oo_instance = ServerProxy(options.uri)
session = oo_instance.openSession('string')
book = oo_instance.openBook(session, options.oofile)

sheet_list = oo_instance.getBookSheets(session, book)
print sheet_list
sheet = sheet_list[0]
sheet = 0
print "sheet ", sheet

#print "preview ", oo_instance.getSheetPreview(session, book, sheet)

#print "setting cell 1,1 to random value"
#oo_instance.setCell(session, book, sheet, 1, 1, randint(1, 65535))

pkt_counter = 0
oo_tcp_ports = get_list_of_tcp_ports()
oo_words = get_list_of_words()

q.set_queue_maxlen(5000)

print "trying to run"
try:
        q.try_run()
except KeyboardInterrupt, e:
        print "interrupted"


print "unbind"
q.unbind(AF_INET)

print "close"
q.close()


