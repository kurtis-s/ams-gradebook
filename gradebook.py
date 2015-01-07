import argparse
import gspread
import httplib2
import fileinput

from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client import tools

class Authorizor(object):
    SCOPE = ['https://spreadsheets.google.com/feeds', 'https://docs.google.com/feeds']
    CLIENT_SECRETS_JSON_PATH = 'client_secrets.json'
    CREDENTIALS_FILE_PATH = 'credentials_file'

    def __init__(self):
        storage = Storage(self.CREDENTIALS_FILE_PATH)
        self.storage = storage
        self.credentials = storage.get()

    def _refresh_token(self):
        if self.credentials.access_token_expired:
            http = httplib2.Http()
            self.credentials.refresh(http)
            self.storage.put(self.credentials)

    def _authorize_application(self):
        parser = argparse.ArgumentParser(parents=[tools.argparser])
        flags = parser.parse_args()
        flow = flow_from_clientsecrets(self.CLIENT_SECRETS_JSON_PATH, scope=self.SCOPE)

        authorized_credentials = tools.run_flow(flow, self.storage, flags)

        return authorized_credentials

    def get_credentials(self):
        """Returns an instantiated Credentials object.  Refreshes the Credential's
        access token if the token is expired.
        
        If no existing credentials file is found, attempt to authorize the application
        with Google.  After the application has been authorized, the new Credential is
        stored in CREDENTIALS_FILE_PATH.
        """
        #storage.get() returns None if there is no credentials file
        if self.credentials is None:
            credentials = self._authorize_application(storage)

        self._refresh_token()

        return self.credentials

class Grader(object):
    def __init__(self, worksheet, grade_column_header, first_name_column_header="First Name", last_name_column_header="Last Name"):
        self._grade_column_header = grade_column_header
        self._grade_column_index = None

        self.first_name_column_header = first_name_column_header
        self.last_name_column_header = last_name_column_header

        self.worksheet = worksheet

        self.first_name_list = None
        self.last_name_list = None

        #A list to keep track of students that we couldn't find in the spreadsheet
        self.missing_students = []
        #A list to keep track of students whose grades we couldn't merge because multiple matches were found
        self.multiple_match_students = []
        #A list of cells with grades to update
        self.grades = []

    @property
    def grade_column_header(self):
        return self._grade_column_header

    @grade_column_header.setter
    def grade_column_header(self, value):
        self._grade_column_header = value
        self._grade_column_index = self._col_index_by_cell_value(self.grade_column_header)

    @property
    def _assignment_column_index(self):
        if self._grade_column_index is None:
            self._grade_column_index = self._col_index_by_cell_value(self.grade_column_header)
        return self._grade_column_index

    @_assignment_column_index.setter
    def _assignment_column_index(self, value):
        self._grade_column_index = value

    def _col_index_by_cell_value(self, search_string):
        """Returns the column index for a cell with a particular string value

        A LookupError is raised if multiple cells with the search string are found, or
        if no cells with the search string are found
        """
        cell_list = self.worksheet.findall(search_string)
        if len(cell_list) < 1:
            raise LookupError("Failed to find search string: {}".format(search_string))
        if len(cell_list) > 1:
            raise LookupError("Multiple matches found for search string: {}".format(search_string))

        return cell_list[0].col

    def _populate_name_lists(self):
        """Populates lists consisting of the first and last names in the spreadsheet and saves
        them in memory so we don't have to keep on fetching them over and over from Google sheets
        """
        first_name_col_index = self._col_index_by_cell_value(self.first_name_column_header)
        self.first_name_list = [first_name.lower() if (first_name is not None) else ""
                for first_name in self.worksheet.col_values(first_name_col_index)]

        last_name_col_index = self._col_index_by_cell_value(self.last_name_column_header)
        self.last_name_list = [last_name.lower() if (last_name is not None) else ""
                for last_name in self.worksheet.col_values(last_name_col_index)]

    def _get_row_indices_for_name(self, first_initial, last_name):
        """Returns the row index corresponding to a particular student"""
        if (self.first_name_list is None) or (self.last_name_list is None):
            self._populate_name_lists()

        #Find the indices in the first name list with matching first initial
        first_name_indices = [i for i, name in enumerate(self.first_name_list)
                if (len(name) >= 1) and (name.lower()[0] == first_initial.lower())]
        #Find the indices in the last name list that match the submitted last_name
        last_name_indices = [i for i, name in enumerate(self.last_name_list)
                if (name.lower() == last_name.lower())]

        #We want the intersection of indices that match both the first initial and last name
        #We need to add 1 to each index because gspread uses 1 indexing and python uses 0 indexing
        return [(i + 1) for i in list(set(last_name_indices).intersection(first_name_indices))]

    def update_grade(self, first_initial, last_name, score):
        student_row = self._get_row_indices_for_name(first_initial, last_name)
        if len(student_row) == 0:
            self.missing_students.append("{} {}".format(first_initial, last_name))
        elif len(student_row) > 1:
            self.multiple_match_students.append("{} {}".format(first_initial, last_name))
        else:
            cell = self.worksheet.cell(student_row[0], self._assignment_column_index)
            cell.value = score
            self.grades.append(cell)

    def update_grades(self):
        self.worksheet.update_cells(self.grades)

if __name__=='__main__':
    authorizor = Authorizor()
    credentials = authorizor.get_credentials()
    gc = gspread.authorize(credentials)
    wks = gc.open_by_key('1SqrL7FigyTy9jhZ9pEDBYplRQaXnDf_Iz-8-MT1LN7o').sheet1

    grade_inputs = fileinput.input()
    assignment_label = grade_inputs.next().strip()
    g = Grader(wks, assignment_label)
    for line in grade_inputs:
        first_initial, last_name, score = line.split()
        g.update_grade(first_initial, last_name, score)
    g.update_grades()
