#! /usr/bin/env python3

import json

import yaml
import sys
import csv
import os
import logging.config

import googleapiclient.discovery
import googleapiclient.errors
import google_auth_oauthlib.flow
import google.oauth2.credentials
import google.oauth2.service_account
from oauthlib.oauth2 import OAuth2Error

try:
    with open(os.path.dirname(os.path.realpath(__file__)) + "/logging.yml", "r") as stream:
        config = yaml.load(stream, Loader=yaml.FullLoader)
    logging.config.dictConfig(config)
except FileNotFoundError as e:
    logging.basicConfig(
        format="[%(asctime)s.%(msecs)03dZ] [%(levelname)s] [%(module)s] [%(funcName)s] %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S",
        level=logging.INFO
    )

log = logging.root
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SCRIPT_NAME = os.path.basename(sys.argv[0])


def _build():

    log.debug("=>")

    token_file = os.getenv("GOOGLE_SHEETS_TOKEN")

    if token_file == "" or token_file is None:
        log.error("The token file must be set in the GOOGLE_SHEETS_TOKEN env variable")
        exit()

    if not os.path.exists(token_file):
        log.error("Token file [ %s ] does not exist", token_file)
        exit()

    with open(token_file, "r") as fd:
        token = json.load(fd)
        if "type" in token.keys() and token["type"] == "service_account":
            log.info("using service account")
            creds = google.oauth2.service_account.Credentials.from_service_account_file(
                filename=token_file, scopes=SCOPES)
        else:
            log.info("using OAuth user")
            creds = google.oauth2.credentials.Credentials.from_authorized_user_file(
                filename=token_file, scopes=SCOPES)

    return googleapiclient.discovery.build('sheets', 'v4', credentials=creds)


USAGE = "\nUsage:\n\n" + SCRIPT_NAME + """ <action> <args>\n
actions         args
----------      -----------
read            <sheet_id> <tab> <data_range>
write           append|update <sheet_id> <tab> <data_range> <data> [auto|json_file|csv_file]
clear           <sheet_id> <tab> <data_range>
find_cell       <sheet_id> <tab> <look_in> <search_term>
convert         <file_to_convert_to_json.csv>
get_oauth_token <client_secret.json>

current token used: """ + os.getenv("GOOGLE_SHEETS_TOKEN")

def _cli():
    log.debug("=>")

    global SERVICE

    if len(sys.argv) > 1 and sys.argv[1] not in ["convert", "get_oauth_token"]:
        SERVICE = _build()

    if len(sys.argv) == 5 and sys.argv[1] == "read":
        read(*sys.argv[2:])
    elif len(sys.argv) in [7, 8] and sys.argv[1] == "write":
        write(*sys.argv[2:])
    elif len(sys.argv) == 5 and sys.argv[1] == "clear":
        clear(*sys.argv[2:])
    elif len(sys.argv) == 6 and sys.argv[1] == "find_cell":
        find_cell(*sys.argv[2:])
    elif len(sys.argv) == 3 and sys.argv[1] == "convert":
        convert(*sys.argv[2:])
    elif len(sys.argv) == 3 and sys.argv[1] == "get_oauth_token":
        get_oauth_token(*sys.argv[2:])
    else:
        log.error(USAGE)
        return


def read(sheet_id, tab, data_range, print_csv=True):
    log.info("sheet_id [ %s ] tab [ %s ] data_range [ %s ]", sheet_id, tab, data_range)
    try:
        request = SERVICE.spreadsheets().values().get(spreadsheetId=sheet_id, range=tab + "!" + data_range)
        resp = request.execute()
    except googleapiclient.errors.HttpError as err:
        log.error("google http error %s", err, exc_info=False)
        return {"error": err.resp.status}
    if resp is None or "values" not in resp:
        log.error("error reading data from spreadsheet")
        return
    log.debug("response [ %s ]", json.dumps(resp))
    if print_csv:
        data = resp["values"]
        csv_file = csv.writer(sys.stdout, quoting=csv.QUOTE_MINIMAL, dialect="excel")
        for row in data:
            csv_file.writerow(row)
    return resp


def write(mode, sheet_id, tab, data_range, data, data_type="auto"):
    log.info("mode [ %s ] sheet_id [ %s ] tab [ %s ] data_range [ %s ] data [ %s ] data_type [ %s ]",
             mode, sheet_id, tab, data_range, data, data_type)
    try:
        if (data_type == "auto" and data.endswith(".json")) or data_type == "json_file":
            if os.path.exists(data):
                with open(data) as json_data:
                    body = json.load(json_data)
            else:
                log.error("json file [ %s ] does not exist", data)
                return
        elif (data_type == "auto" and data.endswith(".csv")) or data_type == "csv_file":
            json_data = convert(csv_file=data, print_json=False)
            if json_data == "":
                log.error("could not convert csv file [ %s ] to json", data)
                return
            body = json.loads(json_data)
        else:
            body = json.loads(data)
    except json.decoder.JSONDecodeError:
        log.error("can't interpret data as json")
        return

    try:
        if mode == "update":
            request = SERVICE.spreadsheets().values().update(
                spreadsheetId=sheet_id, range=tab + "!" + data_range, body=body, valueInputOption='USER_ENTERED')
        elif mode == "append":
            request = SERVICE.spreadsheets().values().append(
                spreadsheetId=sheet_id, range=tab + "!" + data_range, body=body, valueInputOption='USER_ENTERED')
        else:
            log.error("invalid mode [ %s ]", mode)
            return
        resp = request.execute()
    except googleapiclient.errors.HttpError as err:
        log.error("google http error %s", err, exc_info=False)
        return
    log.debug("response [ %s ]", json.dumps(resp))
    return resp


def clear(sheet_id, tab, data_range):
    log.info("sheet_id [ %s ] tab [ %s ] data_range [ %s ]", sheet_id, tab, data_range)
    try:
        request = SERVICE.spreadsheets().values().clear(spreadsheetId=sheet_id, range=tab + "!" + data_range)
        resp = request.execute()
    except googleapiclient.errors.HttpError as err:
        log.error("google http error %s", err, exc_info=False)
        return
    log.debug("response [ %s ]", json.dumps(resp))
    return resp


def find_cell(sheet_id, tab, look_in, search_term, print_address=True):

    if str(look_in).isnumeric():
        look_in_type = "row"
    else:
        look_in_type = "column"

    search_range = str(look_in) + ":" + str(look_in)

    resp = read(sheet_id=sheet_id, tab=tab, data_range=search_range, print_csv=False)

    if "values" not in resp.keys():
        ret = ""
    else:
        try:
            if look_in_type == "column":
                index = list(map(lambda m: m[0] if len(m) > 0 else "", resp["values"])).index(search_term)
                ret = look_in, index + 1
            else:
                index = resp["values"][0].index(search_term)
                ret = _sheet_numeric_column_to_letter(index + 1), int(look_in)
        except ValueError:
            ret = ""

    log.debug("cell address is [ %s ]", ret)
    if print_address:
        print(ret)
    return ret


def convert(csv_file, print_json=True):
    log.info("converting csv_file [ %s ]", csv_file)
    try:
        with open(csv_file) as csvFile:
            reader = csv.reader(csvFile)
            ret = "{\"values\": ["
            printed = 0
            for row in reader:
                if printed != 0:
                    ret += ","
                ret += str(json.dumps(row))
                printed = 1
            ret += "]}"
    except IOError:
        log.error("can't process file [ %s ]. Does the file exist and is it readable?", csv_file)
        return
    if print_json:
        print(ret)
    return ret


def get_oauth_token(client_secret):
    if not os.path.exists(client_secret):
        log.error("Client secret file [ %s ] does not exist", client_secret)
        return
    try:
        flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(client_secret, SCOPES)
        creds = flow.run_local_server()
        print(creds.to_json())
    except OAuth2Error as e:
        log.error("flow not completed [ %s ]", e)
    except KeyboardInterrupt as e:
        log.error("flow interrupted [ %s ]", e)


def _sheet_numeric_column_to_letter(num):
    if num < 1:
        log.error("column can't be < 1")
        return ""
    first = int((num-1) / 26)
    if first > 26:
        log.error("the spreadsheet is too large")
        return ""
    elif first == 0:
        first_letter = ""
    else:
        first_letter = chr(first + 64)
    second_letter = chr(int((num-1) % 26) + 65)

    return first_letter + second_letter


if __name__ == "__main__":
    _cli()
else:
    SERVICE = _build()
