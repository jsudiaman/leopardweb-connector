import os
import sys
from collections import namedtuple
from typing import List

from selenium import webdriver
from selenium.webdriver.support.select import Select

Event = namedtuple('Event', ['name', 'time', 'days', 'date_range'])


class LeopardWebClient:
    """Connects to LeopardWeb using Selenium WebDriver."""

    def __init__(self, username: str, password: str, browser: str):
        """
        Initialize the LeopardWebClient.

        :param username: LeopardWeb username
        :param password: LeopardWeb password
        """
        # Set instance variables
        self.username = username
        self.password = password

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

        # Determine which web driver to use
        if browser.lower() == 'phantomjs':
            self.driver = webdriver.PhantomJS()
        elif browser.lower() == 'chrome':
            self.driver = webdriver.Chrome()
        else:
            raise ValueError('Unsupported browser: {}'.format(browser))
        self.driver.implicitly_wait(30)

        # Login
        self.driver.get('http://leopardweb.wit.edu/')
        self.driver.find_element_by_id('username').send_keys(self.username)
        self.driver.find_element_by_id('password').send_keys(self.password)
        self.driver.find_element_by_css_selector('input.Resizable').click()

    def schedule(self, term: str) -> List[Event]:
        """
        Get schedule from LeopardWeb. Event is a named tuple with field names as follows:

        name (str) - The name of the course, e.g. "SENIOR PROJECT COMP SCIENC-LAB - COMP 5501 - 06"
        time (str) - The time in which the course takes place, e.g. "8:00 am - 9:50 am"
        days (str) - Days of the week in which the course takes place, e.g. "MWF"
        date_range (str) - Start and end dates for the course, e.g. "May 08, 2017 - Aug 15, 2017"

        :param term: School term (e.g. "Summer 2017")
        :return: List of Events
        """
        # Navigate to Student Detail Schedule
        self.driver.find_element_by_link_text('Student').click()
        self.driver.find_element_by_link_text('Registration').click()
        self.driver.find_element_by_link_text('Student Detail Schedule').click()
        for option in Select(self.driver.find_element_by_id('term_id')).options:
            if term.lower() in option.text.lower():
                Select(self.driver.find_element_by_id('term_id')).select_by_visible_text(option.text)
                break
        else:
            raise ValueError('Term "{}" not found'.format(term))
        self.driver.find_element_by_css_selector('div.pagebodydiv > form > input[type="submit"]').click()

        # Parse Student Detail Schedule
        schedule = []
        tables = self.driver.find_elements_by_class_name('datadisplaytable')
        for i in range(0, len(tables), 2):
            t1, t2 = tables[i:i + 2]
            course_name = t1.text.splitlines()[0]
            t2_rows = t2.find_elements_by_tag_name('tr')
            for row in t2_rows[1:]:
                cols = row.find_elements_by_tag_name('td')
                schedule.append(Event(name=course_name, time=cols[1].text, days=cols[2].text, date_range=cols[4].text))
        return schedule

    def shutdown(self) -> None:
        """Shuts down the client."""
        self.driver.quit()
