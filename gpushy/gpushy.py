import argparse
import glob
import gspread
import json
from oauth2client.client import GoogleCredentials
import os
import time


class GPushy():
    def __init__(self, notes_dir, sheet_name, note_keyword):
        try:
            self.sh_key = os.environ['TEMCA_GOOGLE_SPREADSHEET_KEY']
        except KeyError:
            raise Exception("Google spreadsheet key not set!")
        self.notes_locations = notes_dir
        self.parser = TemcaNotes(notes_dir, note_keyword)
        self.sheet_name = sheet_name
        self.current_sheet = None

    def connect(self):
        # Generate credentials
        credentials = GoogleCredentials.get_application_default()
        credentials = credentials.create_scoped(
                ['https://spreadsheets.google.com/feeds'])
        # Authorize
        gc = gspread.authorize(credentials)
        # Open sheet
        sh = gc.open_by_key(self.sh_key)
        sheet_list = sh.worksheets()
        # Set current sheet
        self.current_sheet = [s for s in sheet_list if s.title ==
                              self.sheet_name]
        if len(self.current_sheet) != 1:
            raise Exception("Could not properly set current sheet! {}".format(
                self.current_sheet))
        else:
            self.current_sheet = self.current_sheet[0]
        self.setup()

    def disconnect(self):
        pass

    def connected(self):
        return self.current_sheet is not None

    def setup(self):
        rows = [a for a in self.current_sheet.col_values(1) if
                len(a) > 0]
        self.n_rows = len(rows)
        if self.n_rows > 0:
            self.last_row = rows[-1]
        else:
            self.last_row = None

    def sectionize(self, notes):
        return [
            Section(name=n[0], notes=n[1], t_time=n[2],
                    emlode=n[3].split('/')[2].split('_')[0],
                    tank=n[3].split('/')[2].split('_')[-1],
                    ntiles=n[4], nvetos=n[5], rois=n[6]) for n in notes
        ]

    def grab_all_notes(self):
        notes = self.parser.crawl_notes()
        return self.sectionize(notes)

    def update_cell(self, row, section):
        s = section
        f = [s.name, s.number, s.emlode, s.tank, s.notes,
             s.t_time, s.rois, s.ntiles, s.nvetos, s.veto_percent]
        for i, n in enumerate(f):
            try:
                self.current_sheet.update_cell(row, (i + 1), n)
            except:
                time.sleep(10)
                self.current_sheet.update_cell(row, (i + 1), n)

    def push_sections(self, filled_rows, sections):
        while filled_rows < self.current_sheet.row_count and \
                (filled_rows - self.n_rows) < len(sections):
            current_section = sections[filled_rows - self.n_rows]
            self.update_cell((filled_rows + 1), current_section)
            filled_rows += 1
        if filled_rows >= self.current_sheet.row_count:
            while (filled_rows - self.n_rows) < len(sections):
                sect = sections[filled_rows - self.n_rows]
                self.current_sheet.append_row([
                    sect.name, sect.number, sect.emlode, sect.tank, sect.notes,
                    sect.time, sect.rois, sect.ntiles, sect.nvetos
                ])
                filled_rows += 1

    def initial_push(self):
        if not self.connected():
            self.connect()
        if self.n_rows > 1:
            raise Exception("Spreadsheet already populated! "
                            "Please run update function!")
        sections = self.grab_all_notes()
        self.push_sections(1, sections)

    def update(self, stop_number=None):
        if not self.connected():
            self.connect()
        notes = self.parser.crawl_notes(self.last_row, stop_number=stop_number)
        sections = self.sectionize(notes)
        self.push_sections(self.n_rows, sections)


class TemcaNotes():
    def __init__(self, notes_dir, keyword):
        self.notes_dir = os.path.expanduser(notes_dir)
        self.keyword = keyword

    def parse_note(self, fn):
        print("Parsing: {}".format(fn))
        f = json.load(open(fn, 'r'))
        name = f[u'session'][u'name']
        notes = f[u'session'][u'name']
        try:
            finish = f[u'session'][u'finish']
        except KeyError:
            finish = None
        if finish is not None:
            start = f[u'session'][u'start']
            t_time = float(finish) - float(start)
        else:
            t_time = None
        save_dir = str(f[u'save'][u'directory'])
        try:
            tiles = f[u'session'][u'tiles']
        except KeyError:
            tiles = None
        if tiles is not None:
            vetos = len([n for n in tiles if any(n[u'vetoed'])])
        else:
            vetos = None
        rois = f[u'montage'][u'rois']
        return [
                name, notes, t_time, save_dir,
                tiles, vetos, rois
                ]

    def crawl_notes(self, last_slot=None, stop_number=None):
        print("Crawling: {}".format(self.notes_dir))
        if not os.path.exists(self.notes_dir):
            raise Exception("No such directory! {}".format(self.notes_dir))
        if type(last_slot) == str:
            try:
                last_slot = int(str(last_slot).split('_')[-1])
            except ValueError:
                last_slot = None
        elif type(last_slot) == int:
            pass
        else:
            last_slot = None
        notes, dirs = [], []
        stuff = os.listdir(self.notes_dir)
        stuff.sort()
        for n in stuff:
            if os.path.isdir(os.path.join(self.notes_dir, n)):
                if self.keyword in n:
                    if last_slot is not None:
                        dir_n = int(str(n).split('_')[-1])
                        if dir_n > last_slot:
                            if stop_number is not None and dir_n > stop_number:
                                continue
                            else:
                                dirs.append(n)
                        else:
                            continue
                    else:
                        dirs.append(n)
                else:
                    continue
        for d in dirs:
            search = glob.glob('{}/{}/*_finished.json'.format(
                self.notes_dir, d))
            if len(search) == 1:
                notes.append(self.parse_note(search[-1]))
            elif len(search) > 1:
                print("Multiple jsons found at: {}".format(d))
            else:
                search2 = glob.glob('{}/{}/{}.json'.format(
                    self.notes_dir, d, d))
                if len(search2) == 1:
                    notes.append(self.parse_note(search2[-1]))
                elif len(search) > 1:
                    print("Multiple jsons found at: {}".format(d))
        return notes


class Section():
    def __init__(self, name=None, number=None, emlode=None, tank=None,
                 notes=None, t_time=None, rois=None, ntiles=None,
                 nvetos=None):
        self.name = name
        if number is None and len(self.name) > 0:
            self.number = int(str(self.name.split('_')[-1]))
        elif number is not None:
            self.number = number
        if type(emlode) == int:
            self.emlode = emlode
        elif type(emlode) == str:
            split = emlode.split('emlode')
            if len(split) == 1:
                raise Exception("Improper formatting of emlode field")
            else:
                self.emlode = int(split[-1])
        if type(tank) == int:
            self.tank = tank
        elif type(tank) == str:
            split = tank.split('tank')
            if len(split) == 1:
                raise Exception("Improper formatting of tank field")
            else:
                self.tank = int(split[-1])
        self.notes = notes
        self.t_time = t_time
        self.rois = rois
        if ntiles is not None:
            self.ntiles = len(ntiles)
        else:
            self.ntiles = None
        self.nvetos = nvetos
        if nvetos is None or ntiles is None:
            self.veto_percent = None
        else:
            self.veto_percent = float(nvetos) / float(len(ntiles))


if __name__ == '__main__':
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument('-s', '--source', type=str)
    arg_parser.add_argument('-n', '--sheet_name', type=str)
    arg_parser.add_argument('-i', '--initial', action='store_true')
    arg_parser.add_argument('-u', '--update', action='store_true')
    arg_parser.add_argument('-st', '--stop_number', type=int)
    arg_parser.add_argument('-nk', '--note_keyword', type=str)
    ops = arg_parser.parse_args()

    if ops.source is None or ops.sheet_name is None:
        raise Exception("Please supply a source directory and a "
                        "sheet name to update!")
    if ops.note_keyword is None:
        raise Exception("Please supply a note keyword! This allows the parser"
                        "to find all the notes that contain this keyword")
    pusher = GPushy(ops.source, ops.sheet_name, ops.note_keyword)
    if ops.initial:
        pusher.initial_push()
    if ops.update:
        if ops.stop_number:
            st = ops.stop_number
        else:
            st = None
        pusher.update(stop_number=st)
