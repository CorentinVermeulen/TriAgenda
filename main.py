#! /Users/corentinvrmln/Desktop/Python/Triagenda/venv python3

import datetime as dt
import re
from datetime import timedelta

import gspread
import pandas as pd
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

## Connecting to Google Sheets
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive',
         'https://www.googleapis.com/auth/calendar']

calendar_json = 'triagenda-ff4ff7a817e0.json'
""" calendar_json
{
  "type": "service_account",
  "project_id": "project_name",
  "private_key_id": "****",
  "private_key": "***",
  "client_email": "***",
  "client_id": "***",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "***",
  "universe_domain": "googleapis.com"
}
"""

creds = ServiceAccountCredentials.from_json_keyfile_name(calendar_info, scope)
client = gspread.authorize(creds)
calendarId = CALENDAR_ID
service = build('calendar', 'v3', credentials=creds)


def split(text: str):
    # Définition des motifs regex pour chaque type d'élément
    text = text.replace('min in', 'min')
    pattern_dz = re.compile(r'(\d+)\s*min\s*Z(\d+)')
    pattern_int = re.compile(r'(\d+)\s*×\s*\((.*?)\)')

    elements = []

    while True:
        match_dz = pattern_dz.search(text)
        match_int = pattern_int.search(text)
        if match_dz and match_int:
            if match_dz.start() < match_int.start():
                elements.append(f"{match_dz.group(1)}min Z{match_dz.group(2)}")
                text = pattern_dz.sub('', text, 1)
            else:
                elements.append(f"{match_int.group(1)}×({match_int.group(2)})")
                text = pattern_int.sub('', text, 1)

        elif match_dz:
            elements.append(f"{match_dz.group(1)}min Z{match_dz.group(2)}")
            text = pattern_dz.sub('', text, 1)
        elif match_int:
            elements.append(f"{match_int.group(1)}×({match_int.group(2)})")
            text = pattern_int.sub('', text, 1)
        else:
            break

    return elements


class MySheet:
    concerned_sheet = ['OD1', 'zoneTable']
    date_format = "%Y-%m-%d"
    code_d = {"R": "Run", "C": "Cycling", "N": "Swim"}
    unit_d = {"R": "'/km", "C": "W", }
    sbd_d = {"SBD": "Squat Bench Deadlift", "SB": "Squat Bench", "DB": "Deadlift Bench ", "S": "Squat", "B": "Bench",
             "D": "Deadlift"}

    c_removed = 0
    c_added = 0

    def __init__(self, sheet_name):
        self.sheet = client.open(sheet_name)
        self.make_sheet_data()

    def get_sheet_names(self):
        sheet_names = self.sheet.worksheets()
        return sheet_names

    def _get_sheet_data(self, sheet_name):
        sheet_instance = self.sheet.worksheet(sheet_name)
        records_data = sheet_instance.get_all_records()
        records_df = pd.DataFrame.from_dict(records_data)
        return records_df

    def make_sheet_data(self):
        self.data = {}
        for name in self.concerned_sheet:
            self.data[name] = self._get_sheet_data(name)

    def read_day_prog(self, date):
        prog = self.data['OD1']
        day_prog = prog[prog['Date'] == date]

        sbd = self._check_sbd_day(day_prog)
        car = self._check_cardio_day(day_prog)

        title = f"Prog: {sbd[0] + ('-' if sbd[0] and car[0] else '')} {car[0]}"

        text = ""
        text += sbd[1]
        text += car[1]

        if not sbd[0] and not car[0]:
            title = f"Prog: Rest"
            text = "Rest day"

        if len(text.strip()) > 2:
            self._add_calendar_event(date, title, text)

    def _check_cardio_day(self, day_prog):

        code = day_prog['Code'].values[0]
        if code:
            if code[0] == 'N':
                duration = day_prog['Dur'].values[0]
                text = f"<b>{self.code_d.get(code[0], 'Special ')} Workout ({duration}m):</b>\n"
                desc = day_prog['Desc'].values[0]
                desc = desc.replace('+', '\n')
                desc = desc.replace(')', ')\n')
                desc_format = desc.replace('Cr.', '\tCr.')
                text += desc_format + '\n'
            else:
                duration = day_prog['Dur'].values[0]
                text = f"{self.code_d.get(code[0], 'Special ')} Workout ({code} - {duration}min.):\n"
                desc = day_prog['Desc'].values[0] or " "
                desc = re.sub(r'\n', ' ', desc)
                desc = re.sub(r'\s+', ' ', desc)
                desc = split(desc.replace('\n', ''))
                desc_format = '\n'.join(f'\t{element}' for element in desc)
                desc_zones = self._insert_zones(desc_format, code)
                text += desc_zones + '\n\n'
                text += self._add_zone_intervals(desc_format, code)


            return (code, text)
        else:
            return ('', '')

    def _check_sbd_day(self, day_prog):
        code = day_prog['Power'].values[0]
        if code:
            return (code, f"{self.sbd_d[code]} Workout (~ 2h30)\n\n")
        else:
            return ('', '')

    def _insert_zones(self, desc: str, code: str):
        if code[0] not in "RC" or desc == "":
            return desc

        zones = ["Z1", "Z2", "Z3", "Z4", "Z5", "ZX", "ZY"]
        for zi in zones:
            zonecode = f"{code[0]}{zi}"
            middleval = self.data['zoneTable']['MiddleVal'][self.data['zoneTable']['ZoneCode'] == zonecode].values[0]
            desc = desc.replace(zi, f"{zi}@{middleval}")

        return desc

    def _add_zone_intervals(self, desc: str, code: str):
        txt = ""
        zones = ["Z1", "Z2", "Z3", "Z4", "Z5"]
        for zi in zones:
            if zi in desc:
                zonecode = f"{code[0]}{zi}"
                lowVal = self.data['zoneTable']['LowVal'][self.data['zoneTable']['ZoneCode'] == zonecode].values[0]
                highVal = self.data['zoneTable']['HighVal'][self.data['zoneTable']['ZoneCode'] == zonecode].values[0]
                txt += f"{zi} [{lowVal} - {highVal}] {self.unit_d[code[0]]}\n"
        return txt
    def _make_date_free(self, date):
        start_time = dt.datetime.strptime(date, sheet.date_format)
        end_time = start_time + timedelta(days=1)
        start_time_str_iso = start_time.isoformat() + 'Z'
        end_time_str_iso = end_time.isoformat() + 'Z'

        events_result = service.events().list(
            calendarId=calendarId,
            timeMin=start_time_str_iso,
            timeMax=end_time_str_iso,
            singleEvents=True
        ).execute()
        events = events_result.get('items', [])
        if events:
            for event in events:
                service.events().delete(calendarId=calendarId, eventId=event['id']).execute()
                self.c_removed += 1

    def _add_calendar_event(self, date, title, desc):
        self._make_date_free(date)
        start_time = dt.datetime.strptime(date, sheet.date_format)
        end_time = start_time + timedelta(days=1)
        start_time_str = start_time.strftime(self.date_format)
        end_time_str = end_time.strftime(self.date_format)
        event = {
            'summary': title,
            'description': desc,
            'start': {
                'date': start_time_str,
                'time': 'Europe/Paris',  # Remplacez par le fuseau horaire approprié
            },
            'end': {
                'date': end_time_str,
                'timeZone': 'Europe/Paris',  # Remplacez par le fuseau horaire approprié
            },
        }

        event = service.events().insert(calendarId=calendarId, body=event).execute()
        self.c_added += 1

    def add_week_calendar_event(self):
        today = dt.date.today()
        today_str = today.strftime(self.date_format)
        print(f"Synch for {today_str} until {(today + timedelta(days=7)).strftime(self.date_format)}")
        for day in self.data['OD1']['Date']:
            day_dt = dt.datetime.strptime(day, sheet.date_format).date()
            if today <= day_dt <= today + timedelta(days=7):
                self.read_day_prog(day)

        print(f"Synch done: Added {self.c_added} events {('and removed '+str(self.c_removed)+' events' )if self.c_removed else ''}")



if __name__ == '__main__':
    sheet = MySheet('prog Triathlon')
    sheet.add_week_calendar_event()