#!/usr/bin/python
#
#
# resources: see http://wiki.services.openoffice.org/wiki/Python
#

from xmlrpclib import ServerProxy
from random import randint
import nfqueue
from dpkt import ip
from socket import AF_INET, AF_INET6, inet_ntoa
import sys

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

def cb(i,payload):
        global pkt_counter, oo_tcp_ports
        # every n packets, re-read list of ports
        if pkt_counter % 50 == 0:
            print "re-reading list of TCP ports"
            oo_tcp_ports = get_list_of_tcp_ports()
        pkt_counter += 1
        #print "payload len ", payload.get_length()
        data = payload.get_data()
        pkt = ip.IP(data)
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
                        payload.set_verdict(nfqueue.NF_ACCEPT)
                    else:
                        print "Port %d dropped" % dport
                        update_stats_for_port(dport, pkt, 0)
                        payload.set_verdict(nfqueue.NF_DROP)
                else:
                    print "Port %d not known, so dropped" % dport
                    payload.set_verdict(nfqueue.NF_DROP)
        payload.set_verdict(nfqueue.NF_ACCEPT)

        sys.stdout.flush()
        return 1



q = nfqueue.queue()
q.set_callback(cb)
q.fast_open(0, AF_INET)

oo_instance = ServerProxy('http://localhost:8000')

session = oo_instance.openSession('string')
book = oo_instance.openBook(session, '/home/pollux/oowall.ods')

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


