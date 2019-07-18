__author__ = 'antonio franco'

'''
Copyright (C) 2019  Antonio Franco (antonio_franco@live.it)
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
'''

from ics import Calendar, Event
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P, A
from odf.style import Style, TableCellProperties


class IcsReporter(object):
    """
    Simple class that reads an iCal ics file and provides reports both in csv and odf formats.
    """
    def __init__(self, ics_file: str) -> object:
        """
        Initialize the class with the ics file specified in the path ics_file
        :param ics_file: path to the ics file
        """
        super().__init__()
        self.my_Calendar = Calendar(open(ics_file, "r"))

    def write_report_csv(self, out_file: str, sepz: str = ";") -> None:
        """
        Writes a report of all the events in the ics file as a csv, with separator sepz
        :param out_file (string): path for the output csv file
        :param sepz (string): separator (default ';')
        """
        f = open(out_file, "w+")
        f.write(self.__get_header__(sepz))
        for ev in self.my_Calendar.events:
            f.write(self.__get_csv_line(ev, sepz))

        f.close()

    def __get_csv_line(self, ev: Event, sepz: str) -> str:
        """
        creates a csv line for the event "ev"
        :param ev: ics event
        :param sepz: separator as string
        :return: string formatted for the csv
        """
        (title, begin, end, duration, uid, description, created, location, url, transparent) = self.__get_fields__(ev)
        ret_str = self.__surr_apices__(title, True) + sepz
        ret_str += self.__surr_apices__(begin.format('YYYY-MM-DD HH:mm:ss ZZ'), False) + sepz
        ret_str += self.__surr_apices__(end.format('YYYY-MM-DD HH:mm:ss ZZ'), False) + sepz
        ret_str += self.__surr_apices__(str(duration), False) + sepz
        ret_str += self.__surr_apices__(description, True) + sepz
        ret_str += self.__surr_apices__(created.format('YYYY-MM-DD HH:mm:ss ZZ'), False) + sepz
        ret_str += self.__surr_apices__(location, False) + sepz
        ret_str += self.__surr_apices__(url, False) + "\n"

        return ret_str

    def __surr_apices__(self, my_string: str, clean_it: bool = False, sepz: str = ";") -> str:
        """
        Sorrounds the string with apices
        :param my_string: string to surround with apices
        :param clean_it (boolean): if true removes all the new line, separators, and apices
        :param sepz: separator as string
        :return: string surrounded with apices
        """
        if my_string is not None:
            if clean_it:
                my_string = my_string.replace("\n", " ")
                my_string = my_string.replace('"', " ")
                my_string = my_string.replace(sepz, ".")
            return "\"" + my_string + "\""
        else:
            return "\"NA\""

    def __get_fields__(self, ev: Event) -> tuple:
        """
        returns the following fields from the event: title, begin, end, duration, uid, description, created, location, url, transparent
        :param ev: event
        :return: tuple of strings
        """
        title = ev.name
        begin = ev.begin
        end = ev.end
        duration = ev.duration
        uid = ev.uid
        description = ev.description
        created = ev.created
        location = ev.location
        url = ev.url
        transparent = ev.transparent

        return title, begin, end, duration, uid, description, created, location, url, transparent

    def __get_header__(self, sepz: str) -> str:
        """
        Returns the header for the csv file with separator sepz
        :param sepz: separator
        :return: header as string
        """
        headr = ""
        headr += "title" + sepz
        headr += "begin" + sepz
        headr += "end" + sepz
        headr += "duration" + sepz
        headr += "description" + sepz
        headr += "created" + sepz
        headr += "location" + sepz
        headr += "url" + "\n"

        return headr

    def write_report_ods(self, out_file: str, header_color: str = "#66ffff") -> None:
        """
        Writes a report of all the events in the ics file as an ods. Is possible to optionally specify a background color for the header
        :param out_file (string): path for the output ods file
        :param header_color (string): background color for the header as hexadecimal (e.g. "#66ffff")
        """
        textdoc = OpenDocumentSpreadsheet()

        table = Table(name="Events")

        # Header
        ceX = Style(name="ceX", family="table-cell")
        ceX.addElement(TableCellProperties(backgroundcolor=header_color, border="0.74pt solid #000000"))
        textdoc.styles.addElement(ceX)

        tr = TableRow()
        table.addElement(tr)
        cell = TableCell(valuetype="string", stylename="ceX")
        cell.addElement(P(text="title"))
        tr.addElement(cell)
        cell = TableCell(valuetype="string", stylename="ceX")
        cell.addElement(P(text="begin"))
        tr.addElement(cell)
        cell = TableCell(valuetype="string", stylename="ceX")
        cell.addElement(P(text="end"))
        tr.addElement(cell)
        cell = TableCell(valuetype="string", stylename="ceX")
        cell.addElement(P(text="duration"))
        tr.addElement(cell)
        cell = TableCell(valuetype="string", stylename="ceX")
        cell.addElement(P(text="description"))
        tr.addElement(cell)
        cell = TableCell(valuetype="string", stylename="ceX")
        cell.addElement(P(text="created"))
        tr.addElement(cell)
        cell = TableCell(valuetype="string", stylename="ceX")
        cell.addElement(P(text="location"))
        tr.addElement(cell)
        cell = TableCell(valuetype="string", stylename="ceX")
        cell.addElement(P(text="url"))
        tr.addElement(cell)

        # Content
        for ev in self.my_Calendar.events:
            tr = TableRow()
            (title, begin, end, duration, uid, description, created, location, url, transparent) = self.__get_fields__(
                ev)
            table.addElement(tr)
            cell = TableCell(valuetype="string")
            cell.addElement(P(text=title))
            tr.addElement(cell)
            cell = TableCell()
            cell.addElement(P(text=begin.format('YYYY-MM-DD HH:mm:ss ZZ')))
            tr.addElement(cell)
            cell = TableCell()
            cell.addElement(P(text=end.format('YYYY-MM-DD HH:mm:ss ZZ')))
            tr.addElement(cell)
            cell = TableCell()
            cell.addElement(P(text=str(duration)))
            tr.addElement(cell)
            cell = TableCell(valuetype="string")
            cell.addElement(P(text=description))
            tr.addElement(cell)
            cell = TableCell()
            cell.addElement(P(text=created.format('YYYY-MM-DD HH:mm:ss ZZ')))
            tr.addElement(cell)
            cell = TableCell(valuetype="string")
            cell.addElement(P(text=location))
            tr.addElement(cell)
            cell = TableCell()
            link = A(type="simple", href=url, text=url)
            p = P()
            p.addElement(link)
            cell.addElement(p)
            tr.addElement(cell)

        textdoc.spreadsheet.addElement(table)
        textdoc.save(out_file)


if __name__ == "__main__":
    ics_root = "example"

    C = IcsReporter(ics_root + ".ics")
    C.write_report_csv(ics_root + ".csv")
    C.write_report_ods(ics_root + ".ods")
