import csv
import re
from copy import copy
from datetime import datetime

from openpyxl import load_workbook

TIME_MATCHER = '[0-2][0-9]:[0-5][0-9]'
DATE_MATCHER = '[0-9]{2,4}[\-/.]{1}[0-9]{2}[\-/.]{1}[0-9]{2,4}'
POTENTIAL_DATE_FORMATS = [
    '%d',  # zero padded day
    '%m',  # zero padded month
    '%y',  # two digit year
    '%Y'  # four digit year
]
FORCE_FORMAT = 50  # If format isn't found after this # of records, we will use the best guess
BULK_COUNT = 1000  # The number of records to insert into the database at one time

# Data types
INTEGER = 'integer'
FLOAT = 'float'
STRING = 'varchar'
DATE = 'date'
TIME = 'time'
DATETIME = 'datetime'

CONVERSION_FUNCTIONS = {
    INTEGER: lambda x: int(x),
    FLOAT: lambda x: float(x),
    STRING: lambda x: str(x),
    DATE: lambda x, pattern: datetime.strptime(x, pattern).date(),
    TIME: lambda x, pattern: datetime.strptime(x, pattern).time(),
    DATETIME: lambda x, pattern: datetime.strptime(x, pattern)
}


def readExcelFile(filepath, tabName=None, rowLimit=None, hasHeaders=True):
    """
    This function assumes that data is formatted properly in Excel:
        - The first cell (A1) should contain the start of the data or header
        - There should be no blank columns

    :param filepath: The absolute path to the Excel file
    :param tabName: The name of the tab that should be processed. Only needed if you wish to process a sheet other than the active one.
    :param rowLimit: The number of rows of data that should be loaded into Python. Leave blank if you want to load all data.
    :param hasHeaders: False indicates that there is no header row.
    :return:
    """
    workbook = load_workbook(filepath, data_only=True)
    worksheet = workbook[tabName] if tabName else workbook.active
    columnData = []
    headers = {}  # Keep an dictionary of headerName: idx for easy lookup

    for idx, cell in enumerate(worksheet.rows[0]):
        headerName = cell.value if hasHeaders else None
        if hasHeaders:
            headers[headerName] = idx
        columnData.append(Column(headerName))

    dataRowStart = 1 if hasHeaders else 0

    for row in worksheet.rows[dataRowStart:]:
        for idx, cell in enumerate(row):
            columnData[idx].processValue(cell.value)

        if rowLimit and cell.row == rowLimit:
            break

    for column in columnData:
        column.fixUpRawValues()

    return columnData


def readCSVFile(filepath, rowLimit=None, hasHeaders=True):
    columnData = []
    headers = {}
    with open(filepath, 'rU') as thisFile:
        reader = csv.reader(thisFile)

        for idx, value in enumerate(reader.next()):
            headerName = re.sub(r'[ ]', '', value) if hasHeaders else None
            columnData.append(Column(headerName))
            if hasHeaders:
                headers[headerName] = idx
            else:
                columnData[idx].processValue(value)

        for idx, row in enumerate(reader):
            if rowLimit and idx == rowLimit:
                break
            for nIdx, value in enumerate(row):
                columnData[nIdx].processValue(value)

        for column in columnData:
            column.fixUpRawValues()

    return columnData


class Column(object):
    def __init__(self, headerName, stringToNone=[]):
        self.headerName = headerName
        self.values = []
        self.stringToNone = ['none', 'null', ''] + stringToNone
        self.dataType = None
        self.potentialDatePatterns = []
        self.timeSampleCount = 0
        self.timeFormat = None
        self.dateFormat = None
        self.pattern = None
        self.rawValueIndex = -1  # Set index of values that are stored so we can go back and fix them up once we figure out the format
        self.conversionFunction = None
        self.charLength = 0

    def processValue(self, value):
        value = self.stripValue(value)
        if value is None:
            self.values.append(value)
        else:
            if not self.dataType:
                self.setDataType(value)
            if not self.conversionFunction:
                self.setPattern(value)
            if self.dataType == DATE:
                value = self.normalizeDateValue(value)
            elif self.dataType == TIME:
                value = self.normalizeTimeValue(value)
            elif self.dataType == DATETIME:
                value = self.normalizeDateTimeValue(value)
            if not self.conversionFunction:
                self.rawValueIndex += 1
                self.values.append(value)
            elif self.dataType in [DATETIME, DATE, TIME]:
                try:
                    self.values.append(self.conversionFunction(value, self.pattern))
                except ValueError as e:
                    print('Error occured on header %s with value %s and pattern %s') % (
                        self.headerName, value, self.pattern)
                    raise ValueError(e)
            else:
                try:
                    self.values.append(self.conversionFunction(value))
                except ValueError as e:
                    print('Error occured on header %s with value %s') % (self.headerName, value)
                    raise ValueError(e)

    def fixUpRawValues(self):
        if self.rawValueIndex == -1:
            return
        if not self.conversionFunction:
            raise Exception('Unable to complete conversion. No adequate conversion format found.')
        for i in range(self.rawValueIndex):
            if self.values[i] is None:
                continue
            elif self.dataType in [DATETIME, DATE, TIME]:
                self.values[i] = self.conversionFunction(self.values[i], self.pattern)
            else:
                self.values[i] = self.conversionFunction(self.values[i])

    def stripValue(self, value):
        if value is not None and isinstance(value, str):
            value = value.strip()
            for pat in self.stringToNone:
                if value.lower() == pat.lower():
                    value = None
                    break
            if value is not None:
                self.charLength = max(self.charLength, len(value))
        return value

    def setDataType(self, value):
        if value is None:
            return
        if self.isNumber(value):
            self.dataType = FLOAT if self.isDouble(value) else INTEGER
            self.conversionFunction = CONVERSION_FUNCTIONS[self.dataType]
            return
        hasTime = self.hasTime(value)
        hasDate = self.hasDate(value)
        if hasDate and hasTime:
            self.dataType = DATETIME
        elif hasDate:
            self.dataType = DATE
        elif hasTime:
            self.dataType = TIME
        else:
            self.dataType = STRING
            self.conversionFunction = CONVERSION_FUNCTIONS[self.dataType]

    def isNumber(self, value):
        try:
            int(value)
            return True
        except:
            return False

    def isDouble(self, value):
        return float(value) % 1 != 0

    def hasTime(self, value):
        if re.search(TIME_MATCHER, value):
            return True
        return False

    def hasDate(self, value):
        if re.search(DATE_MATCHER, value):
            return True
        return False

    def setPattern(self, value):
        parts = value.split(' ')
        if self.hasTime and not self.timeFormat:
            self.setTimePattern(parts[1], len(parts) > 2)
        if self.hasDate and not self.dateFormat:
            self.setDatePattern(parts[0])
        if self.dataType == DATETIME:
            if self.timeFormat and self.dateFormat:
                self.pattern = self.dateFormat + ' ' + self.timeFormat
                self.conversionFunction = CONVERSION_FUNCTIONS[self.dataType]
        elif (self.dataType == TIME and self.timeFormat) or (self.dataType == DATE and self.dateFormat):
            self.pattern = self.timeFormat or self.dateFormat
            self.conversionFunction = CONVERSION_FUNCTIONS[self.dataType]

    def normalizeDateTimeValue(self, value):
        return self.normalizeDateValue(self.normalizeTimeValue(value))

    def normalizeDateValue(self, value):
        return re.sub(r'[/]', '-', value)

    def normalizeTimeValue(self, value):
        return re.sub(r'[.]', ':', value)

    def setTimePattern(self, value, hasPM):
        self.timeSampleCount += 1
        timeFormatFound = False
        timeParts = self.normalizeTimeValue(value).split(':')
        pattern = '%I'
        if int(timeParts[0]) > 12:
            pattern = '%H'
            timeFormatFound = True
        pattern += ':%M'
        if len(timeParts) > 2:
            pattern += ':%S'
        if len(timeParts) > 3:
            pattern += '.%f'
        if hasPM:
            pattern += ' %p'
        if timeFormatFound or self.timeSampleCount == FORCE_FORMAT:
            self.timeFormat = pattern

    def setDatePattern(self, value):
        dateParts = [int(x) for x in self.normalizeDateValue(value).split('-')]
        if not self.potentialDatePatterns:
            for _ in dateParts:
                self.potentialDatePatterns.append(copy(POTENTIAL_DATE_FORMATS))
        for idx, part in enumerate(dateParts):
            toRemove = []
            if len(str(part)) == 4:
                self.potentialDatePatterns[idx] = ['%Y']
                toRemove = ['%Y', '%y']
            elif part > 31:
                self.potentialDatePatterns[idx] = ['%y']
                toRemove = ['%Y', '%y']
            elif part > 12:
                self.potentialDatePatterns[idx].remove('%m')
            if len(self.potentialDatePatterns[idx]) == 1:
                toRemove = toRemove or [self.potentialDatePatterns[0]]
                for nIdx in range(len(self.potentialDatePatterns)):
                    if nIdx != idx:
                        for rem in toRemove:
                            self.potentialDatePatterns[nIdx].remove(rem)
        foundPattern = True
        for pat in self.potentialDatePatterns:
            if len(pat) > 1:
                foundPattern = False
                break
        pattern = ''
        if foundPattern:
            for subPat in self.potentialDatePatterns:
                pattern += (subPat[0] + '-')
            self.dateFormat = pattern[:-1]
