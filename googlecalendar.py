"""
Exports LeopardWeb course schedule into Google Calendar.
"""
import argparse
import os
import sys
from collections import OrderedDict
from getpass import getpass

import arrow
import httplib2
from apiclient import discovery
from datetime import timedelta
from oauth2client import client, tools
from oauth2client.file import Storage
from selenium import webdriver
from selenium.webdriver.support.select import Select
from typing import Dict, List
from tzlocal import get_localzone

import util

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'leopardweb-connector'


def get_credentials() -> client.Credentials:
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'leopardweb-connector.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def lw_schedule(username: str, password: str, semester: str, browser: str) -> List[Dict]:
    """
    Get schedule from LeopardWeb.
    
    :param username: LeopardWeb username
    :param password: LeopardWeb password
    :param semester: School semester
    :param browser:  Web browser to use
    :return: Course schedule (formatted as a list of dicts)
    """
    # Determine which web driver to use
    if browser.lower() == 'phantomjs':
        driver = webdriver.PhantomJS()
    elif browser.lower() == 'chrome':
        driver = webdriver.Chrome()
    else:
        raise ValueError('Unsupported browser: {}'.format(browser))
    driver.implicitly_wait(30)

    try:
        # Login
        driver.get('http://leopardweb.wit.edu/')
        driver.find_element_by_id('username').send_keys(username)
        driver.find_element_by_id('password').send_keys(password)
        driver.find_element_by_css_selector('input.Resizable').click()

        # Navigate to "Student Detail Schedule"
        driver.find_element_by_link_text('Student').click()
        driver.find_element_by_link_text('Registration').click()
        driver.find_element_by_link_text('Student Detail Schedule').click()
        Select(driver.find_element_by_id('term_id')).select_by_visible_text(semester)
        driver.find_element_by_css_selector('div.pagebodydiv > form > input[type="submit"]').click()

        # Parse Student Detail Schedule
        schedule = []
        tables = driver.find_elements_by_class_name('datadisplaytable')
        for i in range(0, len(tables), 2):
            t1, t2 = tables[i:i + 2]
            course_name = t1.text.splitlines()[0]
            t2_rows = t2.find_elements_by_tag_name('tr')
            for row in t2_rows[1:]:
                cols = row.find_elements_by_tag_name('td')
                schedule.append({'name': course_name,
                                 'time': cols[1].text,
                                 'days': cols[2].text,
                                 'date_range': cols[4].text})
        print("LeopardWeb schedule is: {}".format(schedule))
        return schedule

    finally:
        driver.quit()


def gc_migrate(events: List[Dict]) -> None:
    """
    Migrate events to Google Calendar.
    
    :param events: The events to migrate
    """
    # Establish connection to Google Calendar
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)
    local_timezone = get_localzone().zone

    # Define dictionaries
    week_dict = OrderedDict()
    week_dict['M'] = 'MO'
    week_dict['T'] = 'TU'
    week_dict['W'] = 'WE'
    week_dict['R'] = 'TH'
    week_dict['F'] = 'FR'
    week_dict['S'] = 'SA'

    month_dict = {'Jan': 'January',
                  'Feb': 'February',
                  'Mar': 'March',
                  'Apr': 'April',
                  'May': 'May',
                  'Jun': 'June',
                  'Jul': 'July',
                  'Aug': 'August',
                  'Sep': 'September',
                  'Oct': 'October',
                  'Nov': 'November',
                  'Dec': 'December'}

    # Loop through LeopardWeb events
    for e in events:
        # Get start and end times
        start_time, end_time = e['time'].split('-')
        start_time = arrow.get(start_time, 'h:mm a').time().isoformat()
        end_time = arrow.get(end_time, 'h:mm a').time().isoformat()

        # Get days of the week (no classes on Sunday, so ignore that case)
        weekdays = ','.join([week_dict[d] for d in e['days']])
        if not weekdays:
            continue

        # Get date range
        start_date, end_date = [util.dict_replace(d, month_dict) for d in e['date_range'].split('-')]

        # Start date
        week_list = list(week_dict.values())
        start_date = arrow.get(start_date, 'MMMM D, YYYY').date()
        while week_list[start_date.weekday()] not in weekdays:
            start_date += timedelta(days=1)
        start_date = start_date.isoformat()

        # End date
        end_date = arrow.get(end_date, 'MMMM D, YYYY').format('YYYYMMDD') + 'T000000Z'

        # Create request body
        body = {
            'summary': e['name'],
            'start': {
                'dateTime': '{}T{}'.format(start_date, start_time),
                'timeZone': local_timezone,
            },
            'end': {
                'dateTime': '{}T{}'.format(start_date, end_time),
                'timeZone': local_timezone,
            },
            'recurrence': [
                'RRULE:FREQ=WEEKLY;BYDAY={};UNTIL={}'.format(weekdays, end_date)
            ],
        }
        body = service.events().insert(calendarId='primary', body=body).execute()
        print('Event created: ' + body.get('htmlLink'))


def main() -> None:
    """
    Main function.
    """
    # Parse arguments
    ap = argparse.ArgumentParser()
    ap.add_argument('-b', '--browser', default='phantomjs', help='web browser')
    args = vars(ap.parse_args())

    # Get user's OS
    if sys.platform.startswith('linux'):
        _os = 'linux'
    elif sys.platform == 'darwin':
        _os = 'osx'
    elif sys.platform.startswith('win'):
        _os = 'windows'
    else:
        raise OSError('Unsupported OS: {}'.format(sys.platform))

    # Add resources to PATH
    os.environ['PATH'] += os.pathsep + os.path.join(os.path.abspath('resources'), _os)

    # Get credentials
    username = input('LeopardWeb Username: ')
    password = getpass('LeopardWeb Password: ')

    # Get LeopardWeb schedule
    events = lw_schedule(username, password, 'Summer 2017 (View only)', args['browser'])

    # Migrate to Google Calendar
    gc_migrate(events)


if __name__ == '__main__':
    main()
