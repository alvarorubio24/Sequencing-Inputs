import os
import platform
import sys
import threading
import time
from math import nan
from pathlib import Path

import pandas as pd
import PySimpleGUI as sg
import tzlocal  # pyinstaller fix
from loguru import logger

import amzl_requests
import report_functions
from database import DatabaseHandler




logger.add(
    "logs\\report_ui.log",
    rotation="10 MB",
    format="{time:YYYY-MM-DD at HH:mm:ss} | {level} | {message}",
    backtrace=True,
    diagnose=True,
    enqueue=True,
)


# chromedriver_path = "chromedriver.exe"
# options = webdriver.ChromeOptions()
# GLOBALS
userprofile = os.getlogin()
threads = []
forecasts = {}
checklists = []


def construct_tab_layout(title: str):  # TODO
    global checklists
    induct_finish = nan
    for checklist in checklists:
        if f"{checklist['node']}-{checklist['cycle']}" == title:
            induct_finish = checklist["plan_sequencing_time"]

    return [
        [
            sg.Text(
                f"{title} Sequencing Report\nFields marked with * are required",
                justification="center",
                size=(55, 2),
            )
        ],
        [sg.Check("Cycle not run", k=f"{title}_not_run")],
        [
            sg.Text("Planned Sequencing Time:", size=(23, 1)),
            sg.In(
                str(induct_finish)[:5], key=f"{title}_pift", size=(5, 10), disabled=True
            ),
            sg.Text("Request Sequencing Time*:", size=(23, 1)),
            sg.In(
                key=f"{title}_aift",
                size=(5, 10),
            ),
        ],
        [
            sg.Text("Sequencing Start Time*:", size=(23, 1)),
            sg.In(key=f"{title}_sequence_start", size=(5, 10)),
            sg.Text("Sequence Finish Time*:", size=(23, 1)),
            sg.In(key=f"{title}_sequence_finish", size=(5, 10)),
        ],
        [
            sg.Text("Routes Assigned Time*:", size=(23, 1)),
            sg.In(key=f"{title}_assigned", size=(5, 10)),
            sg.Text("DPO files sent at:", size=(23, 1)),
            sg.In(key=f"{title}_dpo", size=(5, 10)),
        ],
        [
            sg.Text("SLA:", justification="r"),
            sg.In("", k=f"{title}_SLA", size=(10, 1), disabled=True),
        ],
        [
            sg.Text("SLA miss reason*:", size=(23, 1), justification="r"),
            sg.In("", k=f"{title}_sla_miss_reason", size=(35, 1)),
        ],
        [
            sg.Text("Forced Induct:", visible=False),
            sg.In(key=f"{title}_forced_induct", size=(8, 1), visible=False),
        ],
        [sg.Text("Forced Induct Reason:", visible=False)],
        [
            sg.Multiline(
                key=f"{title}_forced_induct_comment", size=(53, 3), visible=False
            )
        ],
        [sg.Text(" ", visible=False)],
        [
            sg.Text("Number of Route pre-cuts*:", size=(23, 1), justification="r"),
            sg.In(key=f"{title}_precuts", size=(4, 1)),
        ],
        [
            sg.Text("Number of Route pre-adds*:", size=(23, 1), justification="r"),
            sg.In(key=f"{title}_preadds", size=(4, 1)),
        ],
        [
            sg.Text("Routes on Infinity*:", size=(23, 1), justification="r"),
            sg.In(key=f"{title}_infinity", size=(4, 1)),
        ],
        [
            sg.Text("Number of Route cuts*:", size=(23, 1), justification="r"),
            sg.In(key=f"{title}_cuts", size=(4, 1)),
        ],
        [
            sg.Text("Number of Replans*:", size=(23, 1), justification="r"),
            sg.In(key=f"{title}_replans", size=(4, 1)),
        ],
        [
            sg.Text("Replan reason:", size=(23, 1), justification="r"),
            sg.Combo(
                values=[
                    "Add extras",
                    "Adjust cuts",
                    "Balance clusters",
                    "Correct inputs - manual errors",
                    "Decrease minutes",
                    "Sequenced too early",
                    "Increase minutes",
                    "Precuts",
                    "Remove extras",
                    "Add back precuts",
                    "Sequence system error",
                    "Volume not matching",
                ],
                k=f"{title}_replan_reason",
                size=(35, 1),
                tooltip="In case the reason is not on the list, please type your own",
            ),
        ],
        [
            sg.Text("Scheduled Length*:", size=(23, 1), justification="r"),
            sg.In(k=f"{title}_planned_length", size=(12, 1)),
        ],
        [
            sg.Text("Input Minutes*:", size=(23, 1), justification="r"),
            sg.In(key=f"{title}_minutes", size=(12, 1)),
        ],
        [
            sg.Text("DPO DSP Unplanned*:", size=(23, 1), justification="r"),
            sg.In(key=f"{title}_dpo_dsp_unplanned", size=(12, 1)),
        ],
        [
            sg.Text("DPO Flex Unplanned*:", size=(23, 1), justification="r"),
            sg.In(key=f"{title}_dpo_flex_unplanned", size=(12, 1)),
        ],
        [
            sg.Text("DPO Flex Not Considered*:", size=(23, 1), justification="r"),
            sg.In(key=f"{title}_dpo_flex_not_considered", size=(12, 1)),
        ],
        [sg.Text(" ")],
        [sg.Text("Comment:")],
        [sg.Multiline(key=f"{title}_comment", size=(60, 4))],
        [
            sg.In(disabled=True, key=f"{title}_package_count_path", size=(20, 1)),
            sg.FileBrowse(
                "Package Count File*",
                key=f"{title}_package_count",
                file_types=(("Package Count File", "*.csv"),),
            ),
            sg.In(key=f"{title}_package_count_downloaded", visible=False),
        ],
        # hidden fields
        [
            sg.In(k=f"{title}_scheduler", visible=False),
            sg.In(os.getlogin(), k=f"{title}_report_sender", visible=False),
        ],
    ]


def construct_tabs(ds_list):
    return sg.TabGroup([[sg.Tab(tab, construct_tab_layout(tab)) for tab in ds_list]])


# sg.theme("LightPurple")
sg.theme("LightBrown1")


layout_loading = [
    [sg.Text("Downloading Checklist", k="desc", size=(28, 1), auto_size_text=False)],
    [sg.ProgressBar(5, orientation="horizontal", size=(30, 30), k="bar")],
]


def refresh_ds(filepath: str):
    output = []
    df = pd.read_excel(filepath)
    df["Cycle"] = df["Cycle"].astype(str).str.upper().str.strip("_")
    df.fillna("", inplace=True)
    output = list(df.iloc[:, 1])
    output = set(output)
    output = list(output)
    output.sort()
    return output


def list_ds_report() -> tuple[list, list]:
    global checklists
    allowed_cycles = [f"CYCLE_{i}" for i in range(5)]
    allowed_cycles = [*allowed_cycles, *[f"SPECIAL_HANDLING_{i}" for i in range(10)]]
    allowed_cycles = [*allowed_cycles, *[f"SH_{i}" for i in range(10)]]
    nodes_list = []
    nodes_cycles_list = []
    for checklist in checklists:
        node = checklist["node"]
        cycle = checklist["cycle"]
        if cycle in allowed_cycles:
            nodes_list.append(node)
            nodes_cycles_list.append(f"{node}-{cycle}")
    nodes_list = list(set(nodes_list))
    nodes_cycles_list = list(set(nodes_cycles_list))
    return (
        nodes_cycles_list,
        nodes_list,
    )


def parse_report(ds_list, report_input):
    output = {}
    for ds in ds_list:
        keys = [
            key
            for key in report_input.keys()
            if type(key) is str and key.startswith(ds)
        ]
        ds_len = len(ds) + 1
        output[ds] = {key[ds_len:]: str(report_input[key]).strip(" \n") for key in keys}
    return output


def is_empty(val):
    s = str(val).strip("\n ,.")
    if len(s) > 0:
        return 0
    return 1


def verify_values(values: dict, cycle_list: list):
    output = {}
    for cycle in cycle_list:
        output[cycle] = 0
        if values[f"{cycle}_not_run"]:
            pass
        else:
            for key in [k for k in values if str(k).startswith(cycle)]:
                for required in [
                    "aift",
                    "sequence_start",
                    "sequence_finish",
                    "assigned",
                    "sla_miss_reason",
                    "precuts",
                    "preadds",
                    "infinity",
                    "cuts",
                    "replans",
                    "minutes",
                    "dpo_dsp",
                    "dpo_flex",
                ]:
                    if required in key and "forced_induct" not in key:
                        # print(f"{key}: {values[key]}")
                        output[cycle] += is_empty(values[key])
                            
    return output


def main_script_sequencingreport(ds, ofd_date, do_not_save=""):
    global threads
    global forecasts
    global checklists

    db = DatabaseHandler()
    checklists = db.get_station_checklists([ds])
    try:
        country = checklists[0]["country"]
    except IndexError:
        sg.Popup("Station not found in checklist")
        return None

    if country == "GB":
        country = "UK"

    send_to = db.get_email(ds)

    report_filepath = f"output\\{ofd_date}-{ds}"

    os.makedirs(report_filepath, exist_ok=True)

    nodes_cycles_list, nodes_list = list_ds_report()
    f1 = sg.Frame(
        "",
        border_width=0,
        pad=(0, 0),
        element_justification="left",
        layout=[[sg.Button("Load data")]],
    )
    f3 = sg.Frame(
        "",
        border_width=0,
        element_justification="right",
        layout=[[sg.Button("OK")]],
        pad=(0, 0),
    )
    f2 = sg.Frame(
        "",
        border_width=0,
        element_justification="center",
        pad=(0, 0),
        layout=[
            [
                sg.Text(" ", size=(13, 1), auto_size_text=False),
                sg.Text(
                    " S ",
                    k="siphon_daemon",
                    background_color="red",
                    font="Verdana 8",
                    pad=((5, 0), 5),
                    text_color="white",
                    tooltip="Siphon Forecast Status",
                ),
                sg.Text(
                    " V ",
                    k="volume_daemon",
                    pad=((1, 0), 5),
                    background_color="red",
                    font="Verdana 8",
                    text_color="white",
                    tooltip="Volume Forecast Status",
                ),
                sg.Text(
                    " F ",
                    k="flex_daemon",
                    pad=((1, 0), 5),
                    background_color="red",
                    font="Verdana 8",
                    text_color="white",
                    tooltip="Flex Forecast Status",
                ),
                sg.Text(" ", auto_size_text=False, size=(18, 1)),
            ]
        ],
    )
    layout_main = [[construct_tabs(nodes_cycles_list)], [f1, f2, f3]]

    def fetch_siphon():
        report_functions.download_siphon_data(nodes_list, ofd_date).to_csv(
            f"{report_filepath}\\siphon_data.csv", index=False
        )
        forecasts["siphon"] = "green"

    siphon_daemon = threading.Thread(target=fetch_siphon, daemon=True)
    threads.append(siphon_daemon)

    def fetch_volume():
        report_functions.download_siphon_volume(nodes_list, ofd_date).to_csv(
            f"{report_filepath}\\volume_data.csv", index=False
        )
        forecasts["volume"] = "green"

    volume_daemon = threading.Thread(target=fetch_volume, daemon=True)
    threads.append(volume_daemon)

    def fetch_flex():
        report_functions.download_flex_data(
            nodes_list, country.upper(), ofd_date
        ).to_csv(f"{report_filepath}\\flex_data.csv", index=False)
        forecasts["flex"] = "green"

    flex_daemon = threading.Thread(target=fetch_flex, daemon=True)
    threads.append(flex_daemon)

    window = sg.Window("Sequencing Report", layout=layout_main, finalize=True)
    siphon_daemon.start()
    volume_daemon.start()
    flex_daemon.start()
    while True:

        event, values = window.read(timeout=1000)
        if event == sg.WIN_CLOSED:
            break
        for i in values:
            if str(i).endswith("_not_run"):
                current_tab = str(i).replace("_not_run", "")
                for j in values:
                    # Cycle not run button logic
                    if str(j).startswith(current_tab) and not str(j).endswith(
                        "_not_run"
                    ):
                        try:
                            sla = int(values[f"{current_tab}_SLA"].split(" ")[0])
                        except Exception:
                            sla = 0
                        if "SLA" in j or "pift" in j:
                            pass
                        elif "minutes" in j and not is_empty(values[j]):
                            pass
                        elif "replan_reason" in j and values[j] == "None":
                            pass
                        elif "sla_miss_reason" in j:
                            if sla <= 30:
                                window[f"{current_tab}_sla_miss_reason"].update(
                                    "None", disabled=True
                                )
                            else:
                                if values[f"{current_tab}_sla_miss_reason"] == "None":
                                    window[f"{current_tab}_sla_miss_reason"].update("")
                                window[f"{current_tab}_sla_miss_reason"].update(
                                    disabled=False
                                )

                        else:
                            window[j].update(
                                disabled=values[f"{current_tab.strip('_')}_not_run"]
                            )
                try:
                    start_sla = report_functions.parse_time(
                        values[f"{current_tab}_aift"]
                    )
                    finish_sla = report_functions.parse_time(
                        values[f"{current_tab}_assigned"]
                    )
                    sla = report_functions.calculate_sla(
                        (start_sla, finish_sla), "flag"
                    )
                    window[f"{current_tab}_SLA"].update(sla)
                except Exception:
                    pass

        window["siphon_daemon"].update(background_color=forecasts.get("siphon", "red"))
        window["volume_daemon"].update(background_color=forecasts.get("volume", "red"))
        window["flex_daemon"].update(background_color=forecasts.get("flex", "red"))

        if event == "OK":
            values_check = verify_values(values, nodes_cycles_list)
            if sum(values_check.values()) > 0:
                msg = f"Empty fields detected in {[key for key in values_check if values_check[key]>0]}"
                sg.popup(msg, title="Error", background_color="red")
            else:
                for thread in threads:
                    thread.join(timeout=15)
                cycle_list = list_ds_report()[0]
                df = pd.DataFrame.from_dict(
                    parse_report(cycle_list, values), orient="index"
                )
                summary = report_functions.parse_report_inputs(df)
                summary.to_csv(f"{report_filepath}\\summary.csv")

                # Try to save the report to sharedrive - disk is full error
                if not do_not_save:
                    tries = 0
                    while tries <= 100:
                        try:
                            summary.to_csv(
                                f"\\\\ant\\dept-eu\\TBA\\UK\\Business Analyses\\CentralOPS\\AM Shift\\1. Station Checklists\\1. Checklist Manager\\New Checklist\\reports\\{ofd_date}-{ds}.csv"
                            )
                            if (
                                os.stat(
                                    f"\\\\ant\\dept-eu\\TBA\\UK\\Business Analyses\\CentralOPS\\AM Shift\\1. Station Checklists\\1. Checklist Manager\\New Checklist\\reports\\{ofd_date}-{ds}.csv"
                                ).st_size
                                > 0
                            ):
                                break
                            else:
                                tries += 1
                        except Exception:
                            time.sleep(0.5)
                            tries += 1

                for key in values:
                    if "package_count_path" in str(key):
                        if values[key]:
                            if "Package Count" in values[key]:
                                filename = key[: key.find("_")]
                                report_functions.parse_package_count(
                                    values[key]
                                ).to_csv(
                                    f"{report_filepath}\\{filename}.csv", index=False
                                )

                fields = report_functions.gather_tables(
                    report_filepath, ds, ofd_date=ofd_date
                )
                email_body = report_functions.create_html(fields)
                report_functions.create_email(
                    country,
                    report_functions.prettify_html(email_body),
                    ds,
                    send_to,
                    date=ofd_date,
                )
        if event == "Load data":
            active_tab = values[0]
            station = active_tab.split("-")[0]
            cycle = active_tab.split("-")[1]
            auth = amzl_requests.Auth()
            try:
                rtw = amzl_requests.RtwData(
                    station, cycle, ofd_date=ofd_date, auth=auth
                )
                rtw.get_data()
                inputs = rtw.get_readable_inputs()
                plans = rtw.get_plan_info()
                try:
                    dpo = amzl_requests.DpoData(
                        station, cycle, ofd_date=ofd_date, auth=auth
                    )
                    dpo_values = dpo.get_unplanned_routes()
                except Exception:
                    logger.exception("Unable to retrieve DPO data")
                    dpo_values = {}
                try:
                    inputs_values = report_functions.summarise_inputs(inputs)
                    plans_values = report_functions.summarise_plans(plans, country)
                except Exception:
                    inputs_values, plans_values = {}, {}
            except Exception:
                logger.exception("exception happened")

            try:
                window[f"{active_tab}_replans"].update(value=plans_values["replans"])
                if not plans_values["replans"]:
                    window[f"{active_tab}_replan_reason"].update(
                        value="None", disabled=True
                    )
                window[f"{active_tab}_cuts"].update(value=inputs_values["cuts"])
                window[f"{active_tab}_minutes"].update(
                    value=inputs_values["minutes"], disabled=True
                )
                window[f"{active_tab}_infinity"].update(value=inputs_values["infinity"])
                inputs_values["package_count"].to_csv(
                    f"{report_filepath}\\{station}-{cycle}.csv", index=False
                )
                window[f"{active_tab}_package_count"].update(
                    disabled=True, visible=False
                )
                window[f"{active_tab}_package_count_path"].update(
                    visible=False, value=""
                )

                window[f"{active_tab}_dpo_dsp_unplanned"].update(
                    visible=True,
                    value=dpo_values.get("dsp_unplanned", ""),
                    disabled=True,
                )
                window[f"{active_tab}_dpo_flex_unplanned"].update(
                    visible=True,
                    value=dpo_values.get("flex_unplanned", ""),
                    disabled=True,
                )
                window[f"{active_tab}_dpo_flex_not_considered"].update(
                    visible=True,
                    value=dpo_values.get("flex_not_considered", ""),
                    disabled=True,
                )

                window[f"{active_tab}_sequence_start"].update(
                    value=plans_values["start_time"]
                )
                window[f"{active_tab}_sequence_finish"].update(
                    value=plans_values["finish_time"]
                )
                window[f"{active_tab}_scheduler"].update(plans_values["scheduler"])
                planned_length = report_functions.get_route_length(
                    station, cycle, ofd_date
                )
                if planned_length:
                    window[f"{active_tab}_planned_length"].update(
                        planned_length, disabled=True
                    )
                window.refresh()

            except Exception:
                logger.exception(
                    "Uncaught Exception happened, while downloading data from RTW"
                )

            try:
                p = Path(f"{os.getcwd()}\\output\\{ofd_date}-{ds}")
                report = pd.read_csv(p / "summary.csv")
                cycle_string = active_tab.split("-")
                cycle_string = f"{cycle_string[0]}-C{cycle_string[1]}"
                report = report.loc[report["Unnamed: 0"] == active_tab].fillna("")
                for col in [
                    "aift",
                    "assigned",
                    "dpo",
                    "sla_miss_reason",
                    "precuts",
                    "preadds",
                    "replan_reason",
                    "comment",
                ]:
                    try:
                        existing_value = report[col].values[0]
                        if str(existing_value) in ("0.0", "0") or (
                            existing_value and "nan:nan" not in str(existing_value)
                        ):
                            window[f"{active_tab}_{col}"].update(report[col].values[0])
                    except IndexError:
                        pass
            except FileNotFoundError:
                logger.warning("existing report file not found")







    # try:
    #     try:
    #         arg3 = sys.argv[3]
    #     except IndexError:
    #         arg3 = ""
    #     main(sys.argv[1], sys.argv[2], arg3)
        
    #     # sample function call for debugging purposes
    #     # main("DHP1", "2021-11-15", True)  # noqa:E800

    # except Exception:
    #     logger.exception("Uncaught exception happened")