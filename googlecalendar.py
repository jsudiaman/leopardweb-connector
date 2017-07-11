"""
Imports LeopardWeb course schedule into Google Calendar.
"""
import argparse
import os
import re
from collections import OrderedDict
from datetime import timedelta
from getpass import getpass
from typing import Dict, List

import arrow
import httplib2
from apiclient import discovery
from oauth2client import client, tools
from oauth2client.client import Credentials
from oauth2client.file import Storage
from tzlocal import get_localzone

from leopardweb import LeopardWebClient, Event

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'leopardweb-connector'


def get_credentials() -> Credentials:
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


def import_to_google(events: List[Event]) -> None:
    """
    Import LeopardWeb events into Google Calendar.
    
    :param events: The LeopardWeb events to import
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
    for event in events:
        # Get start and end times
        start_time, end_time = event.time.split('-')
        start_time = arrow.get(start_time, 'h:mm a').time().isoformat()
        end_time = arrow.get(end_time, 'h:mm a').time().isoformat()

        # Get days of the week (no classes on Sunday, so ignore that case)
        weekdays = ','.join([week_dict[d] for d in event.days])
        if not weekdays:
            continue

        # Get date range
        start_date, end_date = [dict_replace(d, month_dict) for d in event.date_range.split('-')]

        # Start date
        week_list = list(week_dict.values()) + ['SU']
        start_date = arrow.get(start_date, 'MMMM D, YYYY').date()
        while week_list[start_date.weekday()] not in weekdays:
            start_date += timedelta(days=1)
        start_date = start_date.isoformat()

        # End date
        end_date = arrow.get(end_date, 'MMMM D, YYYY').format('YYYYMMDD') + 'T235959Z'

        # Create request body
        body = {
            'summary': event.name,
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

        # Execute request
        body = service.events().insert(calendarId='primary', body=body).execute()
        print('Event created: {}'.format(body.get('htmlLink')))


def dict_replace(s: str, d: Dict) -> str:
    """
    Replaces all dictionary keys in a string with their respective dictionary values.

    :param s: The string
    :param d: The dictionary
    :return: The new string
    """
    pattern = re.compile('|'.join(d.keys()))
    return pattern.sub(lambda x: d[x.group()], s)


def main() -> None:
    """Main function."""
    # Parse arguments
    ap = argparse.ArgumentParser()
    ap.add_argument('-u', '--username', help='LeopardWeb username')
    ap.add_argument('-p', '--password', help='LeopardWeb password')
    ap.add_argument('-b', '--browser', default='phantomjs', help='Web browser')
    ap.add_argument('-t', '--term', help='School term, e.g. "Summer 2017"')
    args = vars(ap.parse_args())

    # Get LeopardWeb credentials
    if args['username'] is None:
        args['username'] = input('LeopardWeb Username: ')
    if args['password'] is None:
        args['password'] = getpass('LeopardWeb Password: ')
    if args['term'] is None:
        args['term'] = input('Term: ')

    # Import into Google Calendar
    lw = LeopardWebClient(args['username'], args['password'], args['browser'])
    try:
        events = lw.schedule(args['term'])
        print('LeopardWeb events: {}'.format(events))
        import_to_google(events)
        print('Import complete.')
    finally:
        lw.shutdown()


if __name__ == '__main__':
    main()
