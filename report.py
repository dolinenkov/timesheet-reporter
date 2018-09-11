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


class ReportGenerator:

  def __init__(self, password):
    with open('report-config.json', 'rb') as f:
      options = json.load(f)
      self._office = options['office'] # TODO remove: deparment
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
    email = outlook.CreateItem(0x0) #olMailItem = 0x0
    email.To = self._to
    email.CC = self._cc
    email.Subject = self._subject()
    email.Body = self._body()
    print self._body()
    #email.BodyFormat = 2 # olFormatHTML, 3 = olFormatRichText
    #email.HTMLBody = ''
    return email

  def open_report(self):
    self._create_report().Display(False)

  def send_report(self):
    print 'Function disabled'
    #self._create_report().Send()


if __name__ == '__main__':
  ReportGenerator(getpass('Please enter your JIRA password:')).open_report()
