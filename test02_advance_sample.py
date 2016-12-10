#!/usr/bin/env python
# coding: utf-8
# ex: ./getReportAndMail.py "http://127.0.0.1:5000/embed/query/10/visualization/12?api_key=3123123123" "scott@aaa.com;scott@bbb.com"

import os
import urllib2
import re
import sys


# Global Const
ROBOT_MAIL_FROM = os.environ['ROBOT_MAIL_FROM']
ROBOT_SMTP_HOST = os.environ['ROBOT_SMTP_HOST']
ROBOT_SMTP_PORT = os.environ['ROBOT_SMTP_PORT']
ROBOT_SMTP_USER = os.environ['ROBOT_SMTP_USER']
ROBOT_SMTP_PASSWORD = os.environ['ROBOT_SMTP_PASSWORD']


def get_xlsx_file(target_url):
    # targetUrl = 'http://127.0.0.1:5000/embed/query/10/visualization/12?api_key=4932efdb7142df0d15e85a61562157414814266f'
    tmp = urllib2.urlopen(target_url).read()
    title = re.search('.*?visualization.*?"name": "(?P<title>.*?)"', tmp).group('title').decode('unicode-escape')
    table_index = re.search('.*?visualization.*?"id": (?P<id>.*?),', tmp).group('id').decode('unicode-escape')
    result_index = re.search('.*?query_result.*id": (?P<id>.*?)}', tmp).group('id')
    api_key = re.search('.*?api_key=(?P<api_key>.*)', target_url).group('api_key')
    host = re.search('(?P<host>.*?//.*?)/', url).group('host')
    url = host + '/api/queries/%(table_index)s/results/%(result_index)s.xlsx?api_key=%(api_key)s' % {
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


def sendmail(recievers, target_urls=[]):
    reciever_list = tuple(recievers.split(';'))

    # Content Part
    multipart_msg = MIMEMultipart()
    multipart_msg['Subject'] = "Auto Report System"
    multipart_msg['From'] = ROBOT_MAIL_FROM
    multipart_msg['To'] = 'aa@taisys.com,bb@taisys.com,cc@taisys.com'  # 可以造假
    # multipart_msg['To'] = ','.join(reciever_list)
    multipart_msg.attach(MIMEText('Sent by auto-report-robot.'))

    # Attachment Part
    for url in target_urls:
        obj = get_xlsx_file(url)
        part = MIMEBase('application', "octet-stream")
        part.set_payload(urllib2.urlopen(obj.get('url')).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', ('attachment; filename="%(title)s.xls"' % obj).encode('utf-8'))  # must be utf-8 string
        multipart_msg.attach(part)

    # Do send
    s = smtplib.SMTP(ROBOT_SMTP_HOST, ROBOT_SMTP_PORT)
    server = s
    server.ehlo()
    server.starttls()
    s.set_debuglevel(1)
    s.login(ROBOT_SMTP_USER, ROBOT_SMTP_PASSWORD)
    s.sendmail(ROBOT_MAIL_FROM, reciever_list, multipart_msg.as_string())

def parse_params():
    target_urls = sys.argv[1:-1]
    recievers = sys.argv[-1]
    return {
        'target_urls': target_urls,
        'recievers': recievers
    }


def main():
    obj = parse_params()
    sendmail(recievers=obj.get('recievers'), target_urls=obj.get('target_urls'))


main()
