"""
docx2csv.py parses .docx files to extract artwork information, then attempts to pair them up with image files residing in the same directories via filename similarity. It finally outputs a CSV associating each image file to the data parsed from the .docx sharing its name.

It can be run using: "python docx2csv.py ROOT_PATH", or, alternatively the path will be asked for at runtime.

The script expects the directory structure to only have 1 level of depth, meaning it should be:

root_path
    - directory 1
        - file 1
        - file 2
    - directory 2
        - file 1
        - file 2
--
docx2csv.py analiza los archivos .docx para extraer la información de las ilustraciones, y luego intenta emparejarlos con los archivos de imagen
que residen en los mismos directorios a través de la similitud de los nombres de los archivos. Finalmente, genera un CSV que asocia cada archivo de imagen a los datos analizados del .docx que comparte su nombre.

Se puede ejecutar utilizando: "python docx2csv.py ROOT_PATH", o, alternativamente, se pedirá la ruta en tiempo de ejecución.

El script espera que la estructura de directorios sólo tenga un nivel de profundidad, es decir, que sea:

root_path
    - directory 1
        - file 1
        - file 2
    - directory 2
        - file 1
        - file 2

Authors: marting & tfarienti at Museos Abiertos [https://museoabiertos.org]. 2022-05-01
"""

import os
import re
import sys
import csv
import datetime

import docx

def _get_current_time() -> str:
    """ Get the current date and time as a string """
    now = datetime.datetime.now()
    return f'{now.year}-{now.month}-{now.day} {"{:02d}".format(now.hour)}:' \
           f'{"{:02d}".format(now.minute)}:{"{:02d}".format(now.second)}'


def _get_current_time_for_filename() -> str:
    """ Get the current date and time fixed for usage in filenames """
    return _get_current_time().replace(':', '-')


class Docx2CSV:
    """
    Main container class for the scraper operations.

    Takes in a "root path" that must be comprised of several sub-folders, each of those containing .docx documents
    and .jpg files related to said documents. Outputs a CSV relating the data gathered with these .docx documents
    with each of the associated image files.
    """
    KEY_TO_RE = {
        'Work Agent': re.compile(r'Autor:\W*(?P<match>.+)'),
        'Work Title': re.compile(r'(Título:|TÌtulo:)\W*(?P<match>.+)'),
        'Work ID': re.compile(r'(N° de Inventario:|N∞ de Inventario:)\W*(?P<match>.+)'),
        'Work Type': re.compile(r'(Técnica:|TÈcnica:)\W*(?P<match>.+)'),
        'Work Description': re.compile(r'(Tema:|Tema / Descripción:)\W*(?P<match>.+)'),
        'Work Measurements': re.compile(r'Medida[s]*: \W*(?P<match>.+)'),
        'Work Date': re.compile(r'Año:\W*(?P<match>.+)|Fecha:\W*(?P<match_2>.+)'),
    }

    RE_FOUR_DIGIT_YEAR = re.compile(r'^\d\d\d\d$')
    RE_FOUR_DIGIT_YEAR_NAMED_MONTH = re.compile(r'^(?P<year>\d\d\d\d),* (?P<potential_month>\w+)$')
    RE_NAMED_MONTH_FOUR_DIGIT_YEAR = re.compile(r'^(?P<potential_month>\w+),* (de )*(?P<year>\d\d\d\d)$')
    RE_FOUR_DIGIT_YEAR_NAMED_MONTH_DAY = re.compile(r'^(?P<year>\d\d\d\d),* (?P<potential_month>\w+) (?P<day>\d\d)$')
    RE_DAY_NAMED_MONTH_YEAR = re.compile(r'^(?P<day>\d\d?) de (?P<potential_month>\w+) de (?P<year>\d\d\d\d)$')
    RE_MONTH_SLASH_YEAR = re.compile(r'^(?P<month>\d\d?)/(?P<year>\d+)$')
    RE_YEAR_SLASH_MONTH = re.compile(r'^(?P<month>\d\d?)/(?P<year>\d+)$')

    RE_DD_MM_YYYY = re.compile(r'^(?P<day>\d\d?)-(?P<month>\d\d?)-(?P<year>\d\d\d?\d?)$')

    SPANISH_MONTH_TO_DIGIT_INDEX = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO',
                                    'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']

    def __init__(self, root_path: str):
        self.root_path = root_path
        self.rows: list[dict] = []
        self.alerts: list[str] = []

    @staticmethod
    def _info_msg(msg: str) -> None:
        """ Print information message. """
        print(f'[*] {msg}')

    def _error_msg(self, msg: str, fatal: bool = True) -> None:
        """ Print error message, if fatal is True call sys.exit. """
        print(f'[!] {msg}')
        if fatal:
            sys.exit(0)
        else:
            self.alerts.append(f'[!] {msg}')

    def run(self) -> None:
        """
        Run the scraper.

        1 - Go over each directory on the root path
        2 - For each directory extract data from .docx files
        3 - For each of those .docx files pair up with images
        4 - Write data into CSV
        """
        directories = [d for d in os.listdir(self.root_path) if os.path.isdir(f'{self.root_path}/{d}')]
        self._info_msg(f'Starting scrap on: "{self.root_path}" at {_get_current_time()}, found {len(directories)} '
                       f'directories...')
        for directory in directories:
            if not directory.startswith('.'):  # (.DS_Store)
                self.rows.extend(self._handle_directory(f'{self.root_path}/{directory}'))
        self._info_msg(f'Found {len(self.alerts)} alerts. Logging, please wait...')
        self._info_msg(f'Writing CSV, please wait...')
        self.write_alerts()
        self.write_csv()
        self._info_msg(f'Finished scrap at {_get_current_time()}')

    def _handle_directory(self, path: str) -> list[dict]:
        """
        Handle the directory by separating its contents in .docx and other type of files, going over each .docx,
        attempting to parse the data therein, using ._parse_data(), and then using _find_image_files() to find its
        related images.

        :param path: valid directory for Artist information
        :return: list of dicts relating the .docx data with each image file its associated with
        """
        files = [f for f in os.listdir(path) if not f.startswith('.') and os.path.isfile(os.path.join(path, f))]
        docx_files = [f for f in files if f.lower().endswith('.docx')]
        other_files = list(filter(lambda x: x not in docx_files, files))

        csv_rows = []
        for f in docx_files:
            data = self._parse_data(os.path.join(path, f))
            image_files = self._find_image_files(other_files, f)
            if not image_files:
                self._error_msg(f"No images for: {os.path.join(path, f)}", fatal=False)
            else:
                for imf in image_files:
                    csv_rows.append(data | {'Image File': imf})
                    other_files.remove(imf)

        if other_files:
            self._error_msg(f"Extra images in : {path} -> {', '.join(other_files)}", fatal=False)

        return csv_rows

    def _parse_data(self, file_path: str) -> dict:
        """
        Open the document at file_path using docx.Document, transform it into text, and generate a dict of data
        points by attempting to match the data according to Docx2Csv.KEY_TO_RE.

        :param file_path : valid .docx filepath
        :return : dict with keys as in Docx2CSV.KEY_TO_RE, with matched data.
        """
        data = {}

        d: docx.Document = docx.Document(file_path)
        txt = '\n'.join([p.text for p in d.paragraphs])
        for key in self.KEY_TO_RE:
            match = re.search(self.KEY_TO_RE[key], txt)
            if match:
                if match.group('match'):
                    data[key] = match.group('match')
                else:
                    data[key] = match.group('match_2')
            else:
                data[key] = None
                self._error_msg(f'{key} not found in file: {file_path}', fatal=False)

        return data

    @staticmethod
    def _find_image_files(other_files: list[str], docx_file_name: str) -> list[str]:
        """
        Find image files (.jpg/.jpeg) in other_files, that share full, (or partial in the case of serialized image
        files) filename with docx_file_name

        :param other_files: list of file names in the same directory as docx_file_name
        :param docx_file_name: name of .docx file to check against
        :return: list of matching file names
        """
        normalized_docx_file_name = docx_file_name.split('.')[0].lower()
        image_files = []
        for file in other_files:
            if file.lower().endswith('.jpg') or file.lower().endswith('.jpeg'):
                normalized_potential_image_filename = file.split('.')[0].lower()  # Remove extension and go no-case
                if normalized_potential_image_filename == normalized_docx_file_name:
                    image_files.append(file)
                else:  # Use re to check for NAME-\d+ sequences
                    match = re.search(f'{normalized_docx_file_name}-\d+', normalized_potential_image_filename)
                    if match:
                        image_files.append(file)
        return image_files

    def write_csv(self):
        """ Write rows into CSV """
        headers = ['Work ID', 'File', 'Work Agent', 'Work Title', 'Work Type', 'Work Description', 'Work Measurements',
                   'Work Date']

        output_filename = f'output_{_get_current_time_for_filename()}.csv'
        with open(output_filename, 'w', encoding='utf-8', newline='') as w_file:
            csv_writer = csv.writer(w_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
            csv_writer.writerow(headers)

            for r in self.rows:
                row = [
                    self._remove_end_dots(r['Work ID']),
                    self._remove_end_dots(r['Image File']),
                    self._remove_end_dots(r['Work Agent']),
                    self._remove_end_dots(r['Work Title']),
                    self._remove_end_dots(r['Work Type']),
                    self._remove_end_dots(r['Work Description']),
                    self._remove_end_dots(r['Work Measurements']),
                    self._reformat_dates(self._remove_end_dots(r['Work Date'])),
                ]
                csv_writer.writerow(row)

    def write_alerts(self):
        """ Dump alerts into logfile """
        output_filename = f'alerts_{_get_current_time_for_filename()}.txt'
        with open(output_filename, 'w', encoding='utf-8') as w_file:
            w_file.write('\n'.join(self.alerts))

    def _reformat_dates(self, date_txt: str) -> str:
        """ Attempt to reformat dates into YYYY-MM-DD by handling some special cases """
        def format_stuff(raw_str: str, year: str, month: str, day: str = None, named_month: bool = False) -> str:
            """ Format matched output as YYYY-MM-DD. If named_month is True try to map month into
            SPANISH_MONTH_TO_DIGIT_INDEX , if that fails return raw_str instead of formatting. """
            if named_month:
                if month.upper() in self.SPANISH_MONTH_TO_DIGIT_INDEX:
                    month = Docx2CSV.SPANISH_MONTH_TO_DIGIT_INDEX.index(month.upper()) + 1
                else:
                    return raw_str

            return f'{year if len(year) > 2 else "19" + year}-{"{:02d}".format(int(month))}' \
                   f'{"-{:02d}".format(int(day)) if day else ""}'

        #
        if not isinstance(date_txt, str):
            return date_txt

        date_txt = date_txt.strip()  # Get rid of trailing spaces

        # Get rid of these beforehand
        if 'Sin fecha' in date_txt or 'Circa' in date_txt or "'s" in date_txt or "Posterior" in date_txt \
                or date_txt == 'No presenta' or date_txt == 'No' or date_txt == 'S/F' or date_txt == 'Varias':
            return date_txt
        elif re.search(self.RE_FOUR_DIGIT_YEAR, date_txt):
            return date_txt

        # RE attempts
        # noinspection PyCompatibility
        if match := re.search(self.RE_FOUR_DIGIT_YEAR_NAMED_MONTH, date_txt):
            return format_stuff(date_txt, match.group('year'), match.group('potential_month'), named_month=True)
        elif match := re.search(self.RE_FOUR_DIGIT_YEAR_NAMED_MONTH_DAY, date_txt):
            return format_stuff(date_txt, match.group('year'), match.group('potential_month'),
                                match.group('day'), named_month=True)
        elif match := re.search(self.RE_NAMED_MONTH_FOUR_DIGIT_YEAR, date_txt):
            return format_stuff(date_txt, match.group('year'), match.group('potential_month'), named_month=True)
        elif match := re.search(self.RE_DAY_NAMED_MONTH_YEAR, date_txt):
            return format_stuff(date_txt, match.group('year'), match.group('potential_month'),
                                match.group('day'), named_month=True)
        elif match := re.search(self.RE_MONTH_SLASH_YEAR, date_txt):
            return format_stuff(date_txt, match.group('year'), match.group('month'))
        elif match := re.search(self.RE_DD_MM_YYYY, date_txt):
            return format_stuff(date_txt, match.group('year'), match.group('month'), match.group('day'))

        return date_txt

    @staticmethod
    def _remove_end_dots(txt: str) -> str:
        """ Remove the final dot of some fields in order to clean the field """
        if isinstance(txt, str):
            txt = txt.strip()
            if txt.endswith('.'):
                txt = txt[:-1]
        return txt


if __name__ == '__main__':
    try:
        filepath = sys.argv[1]
    except IndexError:
        filepath = input('Please enter the root directory: ')

    D2C = Docx2CSV(root_path=filepath)
    D2C.run()
