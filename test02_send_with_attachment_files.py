#!/usr/bin/env python
# coding: utf-8

import urllib2
import re
import sys


def get_xlsx_file():
    # targetUrl = 'http://127.0.0.1:5000/embed/query/10/visualization/12?api_key=4932efdb7142df0d15e85a61562157414814266f'
    target_url = sys.argv[1]
    mail_to = sys.argv[2]
    tmp = urllib2.urlopen(target_url).read()
    title = re.search('.*?visualization.*?"name": "(?P<title>.*?)"', tmp).group('title').decode('unicode-escape')
    table_index = re.search('.*?visualization.*?"id": (?P<id>.*?),', tmp).group('id').decode('unicode-escape')
    result_index = re.search('.*?query_result.*id": (?P<id>.*?)}', tmp).group('id')
    api_key = re.search('.*?api_key=(?P<api_key>.*)', target_url).group('api_key')
    url = 'http://192.168.1.170:5000/api/queries/%(table_index)s/results/%(result_index)s.xlsx?api_key=%(api_key)s' % {
        'title': title,
        'table_index': table_index,
        'result_index': result_index,
        'api_key': api_key
    }
    # http://127.0.0.1:5000/api/queries/10/results/20.xlsx?api_key=4932efdb7142df0d15e85a61562157414814266f
    return { 'title': title, 'url': url }


import smtplib
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders


def sendMail():
    tmp = get_xlsx_file()

    user = 'aaa'
    pw   = 'bbb'
    host = 'email-ccaabb.com'
    port = 587
    me   = u'support@mm.com'
    you = ('customer@jj.com',)

    # Content
    body = 'Test #NEW_TEST'
    msg  = ("From: %s\r\nTo: %s\r\n\r\n"
           % (me, ", ".join(you)))
    msg = msg + body

    # Content Part
    multipart_msg = MIMEMultipart()
    multipart_msg.attach(MIMEText(msg))

    # Attachment Part
    part = MIMEBase('application', "octet-stream")
    part.set_payload(urllib2.urlopen(tmp.get('url')).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', ('attachment; filename="%(title)s.xls"' % tmp).encode('utf-8'))  # must be utf-8 string
    multipart_msg.attach(part)

    # Do send mail
    s = smtplib.SMTP(host)
    server = s
    server.ehlo()
    server.starttls()
    s.set_debuglevel(1)
    s.login(user, pw)
    s.sendmail(me, you, multipart_msg.as_string())

sendMail()
