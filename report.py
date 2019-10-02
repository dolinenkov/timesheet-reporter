#!/usr/bin/env python2


"""

daily report sending script

pip dependencies  : pywin32, jira, jinja2, keyring
software needed   : outlook 2013

reference:
https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
https://jira.readthedocs.io/en/master/examples.html
https://pypi.org/project/keyring

"""

import jira
import datetime
import json
import getpass
import win32com.client
import jinja2
import keyring


class Reporter:
    _OPTIONS = u'.report.json'

    _JIRA_QUERY = 'worklogAuthor=currentUser() and worklogDate>=startOfDay(-0d) and worklogDate<=endOfDay(-0d)'

    @staticmethod
    def _read_config():
        with open(Reporter._OPTIONS, 'rb') as file:
            return json.load(file)

    @staticmethod
    def _get_and_save_password(jira_url, jira_login):
        passwd = keyring.get_password(jira_url, jira_login)
        if not passwd:
          passwd = getpass.getpass('Password for {} @{} :>'.format(jira_url, jira_login))
          if passwd:
            keyring.set_password(jira_url, jira_login, passwd)
        return passwd

    @staticmethod
    def _work_log_duration(work_logs, author, date):
        def get_date(w):
            return datetime.datetime.strptime(w.started.split(u'T', 1)[0], u'%Y-%m-%d').date()
        return sum(w.timeSpentSeconds / 3600.0 for w in work_logs if w.author.key == author and get_date(w) == date)

    @staticmethod
    def _work_log_for_issue(work_logs, author, date):
        duration = Reporter._work_log_duration(work_logs, author, date)
        return {
            'today': duration,
            'total': duration
        }

    @staticmethod
    def _work_log_for_project(project):
        return {
            'today': sum(issue['time']['today'] for issue in project['issues']),
            'total': sum(issue['time']['total'] for issue in project['issues'])
        }

    @staticmethod
    def _work_log_for_all_projects(projects):
        return {
            'today': sum(project['time']['today'] for project in projects),
            'total': sum(project['time']['total'] for project in projects)
        }

    @staticmethod
    def _get_timesheet(config, date):
        jira_url = config['jira_url']
        jira_login = config['jira_login']
        jira_password = Reporter._get_and_save_password(jira_url, jira_login)
        j = jira.JIRA(jira_url, basic_auth=(jira_login, jira_password))

        projects = []
        for i in j.search_issues(Reporter._JIRA_QUERY):

            filtered_projets = filter(lambda p: p['name'] == i.fields.project.key, projects)
            if len(filtered_projets) > 0:
                project = filtered_projets[0]
            else:
                project = {
                    'name': i.fields.project.key,
                    'url': '',
                    'issues': []
                }
                projects.append(project)

            project['issues'].append({
                'summary': u'{} - {}'.format(i.key, i.fields.summary),
                'url': '',
                'time': Reporter._work_log_for_issue(j.worklogs(i.key), config['jira_login'], date)
            })

        for project in projects:
            project['time'] = Reporter._work_log_for_project(project)

        return {
            'date': date,
            'time': Reporter._work_log_for_all_projects(projects),
            'projects': projects,
        }

    def _generate_subject(self, options):
        return self._jenv.get_template('email_subject.txt').render(options)

    def _generate_body(self, options):
        return self._jenv.get_template('email_body.html').render(options)

    def _create_mail(self, date):

        class OlItemType:
            olMailItem = 0

            def __init__(self):
                pass

        class OlMailRecipientType:
            olCC = 2
            olTo = 1

            def __init__(self):
                pass

        o = win32com.client.Dispatch('Outlook.Application')
        m = o.CreateItem(OlItemType.olMailItem)

        t = {}
        t.update(self._config)
        t.update(Reporter._get_timesheet(self._config, date))

        m.Subject = self._generate_subject(t)
        m.HTMLBody = self._generate_body(t)
        for r in self._config['mail_to']:
            m.Recipients.Add(r).Type = OlMailRecipientType.olTo
        for r in self._config['mail_cc']:
            m.Recipients.Add(r).Type = OlMailRecipientType.olCC
        m.Recipients.ResolveAll()
        return m

    def __init__(self):
        self._config = Reporter._read_config()
        self._jenv = jinja2.Environment(loader=jinja2.FileSystemLoader(searchpath='./'))

    def _create_report(self):
        return self._create_mail(datetime.date.today())

    def display(self):
        self._create_report().Display(False)


if __name__ == '__main__':
    Reporter().display()
