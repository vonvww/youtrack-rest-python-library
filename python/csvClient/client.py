import sys
import csv
import csvClient

maxInt = sys.maxsize
while True:
    # decrease the maxInt value by factor 10
    # as long as the OverflowError occurs.
    try:
        csv.field_size_limit(maxInt)
        break
    except OverflowError:
        maxInt = int(maxInt/10)


class Client(object):
    def __init__(self, file_path):
        self._file_path = file_path
        self._header = self._read_header()
        self._issues_reader = None

    def _read_header(self):
        header = self._get_reader().next()
        return [field_name
                for field_name in [h.strip() for h in header]
                if len(field_name)]

    def _get_reader(self):
        return csv.reader(open(self._file_path, "r"),
                          delimiter=csvClient.CSV_DELIMITER)

    def get_rows(self):
        reader = self._get_reader()
        for row in reader:
            yield row

    def get_issues(self):
        reader = self._get_reader()
        reader.next()
        header_len = len(self._header)
        for row in reader:
            if not row:
                continue
            issue = {"comments": []}
            for i in range(len(row)):
                value = row[i].strip()
                if len(value):
                    if i < header_len:
                        issue[self._header[i]] = value
                    else:
                        issue["comments"].append(value)
            yield issue

    def get_header(self):
        return self._header

    def reset(self):
        self._issues_reader = self._get_reader()
        self._read_header()
