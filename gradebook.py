import argparse
import fileinput
import gspread
import httplib2
import os
import spreadsheetconfig
import sys

from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client import tools


class Authorizor(object):
    """Gets OAuth2 credentials.
    Looks for OAuth2 parameters in the path specified by the
    client_secrets_json_path (see
    https://developers.google.com/api-client-library/python/guide/aaa_client_secrets).
    Looks for and stores the credentials in the path specified by the
    credentials_file_path attribute.
    """
    scope = ['https://spreadsheets.google.com/feeds', 'https://docs.google.com/feeds']
    client_secrets_json_path = 'client_secrets.json'
    credentials_file_path = 'credentials_file'

    def __init__(self):
        storage = Storage(self.credentials_file_path)
        self.storage = storage
        self.credentials = storage.get()

    def _refresh_token(self):
        if self.credentials.access_token_expired:
            http = httplib2.Http()
            self.credentials.refresh(http)
            self.storage.put(self.credentials)

    def _authorize_application(self):
        parser = argparse.ArgumentParser(parents=[tools.argparser])
        flags = parser.parse_args(args=sys.argv[2:])
        flow = flow_from_clientsecrets(self.client_secrets_json_path,
                                       scope=self.scope)

        authorized_credentials = tools.run_flow(flow, self.storage, flags)

        return authorized_credentials

    def get_credentials(self):
        """Returns an instantiated Credentials object.  Refreshes the
        Credential's access token if the token is expired.

        If no existing credentials file is found, attempts to authorize the
        application with Google.  After the application has been authorized,
        the new Credential is stored in credentials_file_path.
        """
        # storage.get() returns None if there is no credentials file
        if self.credentials is None:
            self.credentials = self._authorize_application()

        self._refresh_token()

        return self.credentials


class Grader(object):
    """Updates grades in the google spreadsheet.

    Args:
        worksheet (gspread.Worksheet): The google spreadsheet gradebook
        grade_column_header (str): The assignment's header in the gradebook
        first_name_column_header (str, optional): The first name column header
            in the gradebook.
        last_name_column_header (str, optional): The last name column header in
            the gradebook.
    """
    def __init__(self, worksheet, grade_column_header, first_name_column_header,
                 last_name_column_header):
        self._grade_column_header = grade_column_header
        self._grade_column_index = None

        self.first_name_column_header = first_name_column_header
        self.last_name_column_header = last_name_column_header

        self.worksheet = worksheet

        self.first_name_list = None
        self.last_name_list = None

        # A list to keep track of students that we couldn't find in the
        # spreadsheet
        self.missing_students = []

        # A list to keep track of students whose grades we couldn't merge
        # because multiple matches were found
        self.multiple_match_students = []

        # A list of cells with grades to update
        self.grades = []

    @property
    def grade_column_header(self):
        """Specifies the assignment's header in the spreadsheet"""
        return self._grade_column_header

    @grade_column_header.setter
    def grade_column_header(self, value):
        self._grade_column_header = value
        self._grade_column_index = self._col_index_by_cell_value(
            self.grade_column_header)

    @property
    def _assignment_column_index(self):
        if self._grade_column_index is None:
            self._grade_column_index = self._col_index_by_cell_value(
                self.grade_column_header)

        return self._grade_column_index

    @_assignment_column_index.setter
    def _assignment_column_index(self, value):
        self._grade_column_index = value

    def _col_index_by_cell_value(self, search_string):
        """Returns the column index for a cell with a particular string value

        Raises:
            LookupError: If multiple cells with the search string are found, or
                no cells with the search string are found
        """
        cell_list = self.worksheet.findall(search_string)
        if len(cell_list) < 1:
            raise LookupError(
                "Failed to find search string: {}"
                .format(search_string))
        if len(cell_list) > 1:
            raise LookupError(
                "Multiple matches found for search string: {}"
                .format(search_string))

        return cell_list[0].col

    def _populate_name_lists(self):
        """Populates lists consisting of the first and last names in the
        spreadsheet and saves them in memory so we don't have to keep on
        fetching them over and over from Google sheets
        """
        first_name_col_index = self._col_index_by_cell_value(
            self.first_name_column_header)
        self.first_name_list = [
            first_name.lower() if (first_name is not None) else ""
            for first_name in self.worksheet.col_values(first_name_col_index)]

        last_name_col_index = self._col_index_by_cell_value(
            self.last_name_column_header)
        self.last_name_list = [
            last_name.lower() if (last_name is not None) else ""
            for last_name in self.worksheet.col_values(last_name_col_index)]

    def _get_row_indices_for_name(self, first_initial, last_name):
        """Returns the row index corresponding to a particular student"""
        if (self.first_name_list is None) or (self.last_name_list is None):
            self._populate_name_lists()

        # Find the indices in the first name list with matching first initial
        first_name_indices = [
            i for i, name in enumerate(self.first_name_list)
            if (len(name) >= 1) and (name.lower()[0] == first_initial.lower())]

        # Find the indices in the last name list matching the submitted last_name
        last_name_indices = [
            i for i, name in enumerate(self.last_name_list)
            if (name.lower() == last_name.lower())]

        # We want the intersection of indices that match both the first initial
        # and last name
        # We need to add 1 to each index because gspread uses 1 indexing and
        # python uses 0 indexing
        return [(i + 1) for i in list(
            set(last_name_indices).intersection(first_name_indices))]

    def add_grade(self, first_initial, last_name, score):
        """Queues up grades that need to be updated in the spreadsheet.  Note
        this method only queues up grades that need to be updated.  The actual
        update is done in batch by calling update_grades"""
        student_row = self._get_row_indices_for_name(first_initial, last_name)
        if len(student_row) == 0:
            self.missing_students.append("{} {} {}\n".format(
                first_initial, last_name, score))
        elif len(student_row) > 1:
            self.multiple_match_students.append("{} {} {}\n".format(
                first_initial, last_name, score))
        else:
            cell = self.worksheet.cell(student_row[0],
                                       self._assignment_column_index)
            cell.value = score
            self.grades.append(cell)

    def update_grades(self):
        """Sends the grades to Google sheets in a batch update"""
        self.worksheet.update_cells(self.grades)

    def save_unmergeable_grades(self, directory):
        """Stores the grades for students that could not be merged to a file
        and prints a warning message in the console

        Args:
            directory (str): Directory where list of students that could not be
                merged should be saved
        """

        missing_student_message = (
            "Failed to merge the following grades because the student's name "
            "was not found in the spreadsheet:\n")
        multiple_students_message = (
            "Failed to merge the following grades because multiple matches "
            "for the student's name were found in the spreadsheet:\n")

        if (len(self.missing_students) > 0) or (len(self.multiple_match_students) > 0):
            with open(os.path.join(directory, 'unmergeable_students.txt'), 'w') as unmergeable_file:
                if(len(self.missing_students) > 0):
                    unmergeable_file.write(missing_student_message)
                    unmergeable_file.writelines(self.missing_students)
                    unmergeable_file.write("\n")

                if(len(self.multiple_match_students) > 0):
                    unmergeable_file.write(multiple_students_message)
                    unmergeable_file.writelines(self.multiple_match_students)

            # Print out the grades that couldn't be merged the console for convenience
            with open(os.path.join(directory, 'unmergeable_students.txt'), 'r') as unmergeable_file:
                print unmergeable_file.read()


def main():
    authorizor = Authorizor()
    credentials = authorizor.get_credentials()
    gc = gspread.authorize(credentials)
    wks = gc.open_by_key(spreadsheetconfig.key).sheet1

    grade_inputs = fileinput.input()
    assignment_label = grade_inputs.next().strip()
    g = Grader(
        wks,
        assignment_label,
        first_name_column_header=spreadsheetconfig.first_name_column_header,
        last_name_column_header=spreadsheetconfig.last_name_column_header)
    for line in grade_inputs:
        # Ignore empty lines and comments
        if (line.strip()) and (line[0] != '#'):
            first_initial, last_name, score = line.split()
            g.add_grade(first_initial, last_name, score)

    g.update_grades()
    g.save_unmergeable_grades(os.path.dirname(sys.argv[1]))

if __name__ == '__main__':
    main()
