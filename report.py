#!/usr/bin/env python2


"""

daily report sending script

pip dependencies  : pywin32, jira
software needed   : outlook 2013

reference:
https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
https://jira.readthedocs.io/en/master/examples.html

"""

import json
from datetime import *
from getpass import getpass
from win32com.client import Dispatch


"""
sending mail
"""


class MailSender(object):

    def _to(self, message):
        return message.get('to', list())

    def _cc(self, message):
        return message.get('cc', list())

    def _subject(self, message):
        return message.get('subject', '')

    def _body(self, message):
        return message.get('body', '')

    def send_mail(self, message):
        raise NotImplementedError()


class OutlookMailSender(MailSender):

    NAME = 'outlook'

    def __init__(self):
        super(OutlookMailSender, self).__init__()

    def send_mail(self, message):

        olMailItem = 0x0
        olFormatPlain = 0 #olFormatHTML = 2, olFormatRichText = 3

        import win32com.client
        outlook = win32com.client.Dispatch('Outlook.Application')
        message = outlook.CreateItem(olMailItem)

        message.To = ';'.join(self._to(message))
        message.CC = ';'.join(self._cc(message))
        message.Subject = self._subject(message)
        message.BodyFormat = olFormatPlain
        message.Body = self._body(message)

        message.Display(False)


class MailSenderFactory(object):

    def __init__(self):
        self._classes = {
            OutlookMailSender.NAME: OutlookMailSender
        }

    def create(self, name, **kwargs):
        c = self._classes.get(name, None)
        return c(kwargs) if c is not None else c


"""
collecting worklogs
"""


class WorklogCollector(object):

    def __init__(self):
        pass

    def collect_worklogs(self):
        raise NotImplementedError()


class JiraWorklogCollector(WorklogCollector):

    NAME = 'jira'

    def __init__(self):
        super(WorklogCollector, self).__init__()

    def collect_worklogs(self):
        return {}


class WorklogCollectorFactory(object):

    def __init__(self):
        self._classes = {
            JiraWorklogCollector.NAME: JiraWorklogCollector
        }

    def create(self, name, **kwargs):
        c = self._classes.get(name, None)
        return c(kwargs) if c is not None else c


"""
"""


class ReportGenerator(object):

    def generate(self):
        pass


class ReportSender(object):

    def __init__(self):
        from argparse import ArgumentParser
        parser = ArgumentParser()

        self._parser = parser

    def run(self, arguments):
        parsed = self._parser.parse_args(arguments)
        worklog_collector = WorklogCollectorFactory().create(parsed.worklog_collector)
        mail_sender = MailSenderFactory().create(parsed.mail_sender)
        mail_sender.send_mail()


if __name__ == '__main__':
    import sys
    ReportSender().run(sys.argv)


class ReportGenerator:

    def __init__(self, password):
        with open('report-config.json', 'rb') as f:
            options = json.load(f)
            self._office = options['office']  # TODO remove: deparment
            self._project = options['project']
            self._to = ';'.join(t for t in options['to'])
            self._cc = ';'.join(t for t in options['cc'])
            self._header = options['header']
            self._footer = options['footer']
            self._login = options['login']
            self._jira = options['jira']
        self._password = password

    def _timesheet(self):
        return '\n{}\n'.format('')

    def _subject(self):
        return '[{}][{}] Daily Report {}'.format(self._office, self._project, datetime.now().strftime('%d.%m.%Y'))

    def _body(self):
        return self._header + self._timesheet() + self._footer

    def _create_report(self):
        outlook = Dispatch("Outlook.Application")
        email = outlook.CreateItem(0x0)  # olMailItem = 0x0
        email.To = self._to
        email.CC = self._cc
        email.Subject = self._subject()
        email.Body = self._body()
        print self._body()
        # email.BodyFormat = 2 # olFormatHTML, 3 = olFormatRichText
        # email.HTMLBody = ''
        return email

    def open_report(self):
        self._create_report().Display(False)

    def send_report(self):
        print 'Function disabled'
        # self._create_report().Send()


if __name__ == '__main__':
    ReportGenerator(getpass('Please enter your JIRA password:')).open_report()
