import argparse
# import ctypes
import os
import platform
import sqlite3
from datetime import datetime

import PySimpleGUI as sg
from loguru import logger

from database import DatabaseHandler




# def make_dpi_aware():
#     if int(platform.release()) >= 8:
#         ctypes.windll.shcore.SetProcessDpiAwareness(True)


logger.add(
    "logs\\checklist_update.log",
    rotation="10 MB",
    enqueue=True,
    format="{time:YYYY-MM-DD at HH:mm:ss} | {level} | {message}",
)


class AbortedException(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)



class ChecklistWindowMultiple:
    def __init__(self, station: str) -> None:
        # make_dpi_aware()
        
        
        sg.theme("LightPurple")
        self.db = DatabaseHandler()
        self.existing_values = self.db.get_station_checklists(
            [station.upper()], active_only=False
        )
        if not self.existing_values:
            sg.PopupError(
                "Station not found in checklist.\nMake sure to use Parent Station while updating multiple checklists"
            )
            raise AbortedException
        self.nodes = [f"{d['node']}-{d['cycle']}" for d in self.existing_values]
        self.window = sg.Window("Checklist Update", self.window_layout())
        #window = Toplevel(root)

    def tab_layout(self, checklist: str) -> list[list]:
        node = checklist.split("-")[0]
        cycle = checklist.split("-")[1]
        label = (20, 1)
        input_size = (60, 1)
        for dct in self.existing_values:
            if f"{dct['node']}-{dct['cycle']}" == checklist:
                ex = dct
        frame1 = [
            [sg.Check("Active?", ex.get("active", True), k=f"{checklist}__active_")],
            [
                sg.Text("Country*:", label),
                sg.Combo(
                    ["AT", "BE", "DE", "ES", "FR", "IE", "IT", "NL", "UK"],
                    default_value=ex.get("country", None),
                    s=(5, 1),
                    readonly=True,
                    k=f"{checklist}__country",
                ),
            ],
            [
                sg.Text("Parent DS*:", label),
                sg.In(
                    ex.get("parent_station", ""), s=(5, 1), k=f"{checklist}__parent_ds"
                ),
            ],
            [
                sg.Text("Type*:", label),
                sg.Combo(
                    ["AMZL", "AMPL", "AMXL"],
                    default_value=ex.get("type", "AMZL"),
                    k=f"{checklist}__node_type",
                    s=(5, 1),
                    readonly=True,
                ),
            ],
            [
                sg.Text("Region*:", label),
                sg.In(ex.get("region", ""), k=f"{checklist}__region", s=input_size),
            ],
            [
                sg.Text("Planned sequencing time\n(station local time)*:", s=(20, 2)),
                sg.In(
                    ex.get("plan_sequencing_time", ""),
                    k=f"{checklist}__plan_sequencing_time",
                    s=input_size,
                ),
            ],
            [
                sg.Text("Flex*:", label),
                sg.In(ex.get("flex", ""), k=f"{checklist}__flex", s=input_size),
            ],
            [
                sg.Text("Autoassign*:", label),
                sg.In(
                    ex.get("autoassign", ""), k=f"{checklist}__autoassign", s=input_size
                ),
            ],
            [
                sg.Text("Upload SCC*:", label),
                sg.In(
                    ex.get("upload_scc", ""), s=input_size, k=f"{checklist}__upload_scc"
                ),
            ],
        ]
        multiline_size = (50, 8)
        frame2 = [
            [sg.Text("Cluster Handling:")],
            [
                sg.Multiline(
                    ex.get("cluster_comments", ""),
                    k=f"{checklist}__cluster_comments_",
                    s=multiline_size,
                )
            ],
            [sg.Text("Remarks:")],
            [
                sg.Multiline(
                    ex.get("comments", ""),
                    k=f"{checklist}__comments_",
                    s=multiline_size,
                )
            ],
        ]
        try:
            last_audit = f"{ex['last_audit'][:19]} by {ex['audited_by']}"
        except KeyError:
            last_audit = ""
        layout = [
            [
                sg.Text("Node:"),
                sg.In(
                    node,
                    k=f"{checklist}__node",
                    disabled=True,
                    s=(6, 1),
                ),
                sg.Text("Cycle:"),
                sg.In(
                    cycle,
                    k=f"{checklist}__cycle",
                    disabled=True,
                    s=(10, 1),
                ),
                sg.Text("Last audit:"),
                sg.In(last_audit, s=(45, 1), disabled=True, k=f"{checklist}__audit_"),
            ],
            [sg.HorizontalSeparator()],
            [
                sg.Frame("", layout=frame1, border_width=0),
                sg.VerticalSeparator(),
                sg.Frame("", layout=frame2, border_width=0),
            ],
        ]

        # return sg.Frame("", layout=layout, border_width=0)
        return layout

    def window_layout(self) -> list[list[sg.Element]]:
        tabs = [[sg.Tab(n, self.tab_layout(n)) for n in self.nodes]]
        layout = [
            [sg.TabGroup(tabs, k="active_tab_")],
            [sg.HorizontalSeparator()],
            [sg.Cancel(), sg.Submit()],
        ]
        return layout

    def output_to_json(self, values: dict) -> list[dict]:
        out = []
        keys = [
            "active_",
            "autoassign",
            "cluster_comments_",
            "comments_",
            "country",
            "cycle",
            "flex",
            "node",
            "node_type",
            "parent_ds",
            "plan_sequencing_time",
            "region",
            "upload_scc",
        ]
        for node in self.nodes:
            dct = {}

            for key in keys:
                dct[f"{key}"] = values[f"{node}__{key}"]
            out.append(dct)
        return out

    def count_changes(self, values: list[dict]) -> dict:
        out = {}
        to_strip = "\n\t ,./\\"
        for node in self.nodes:
            out[node] = 0
            for d in self.existing_values:
                if f"{d['node']}-{d['cycle']}" == node:
                    old = d
            for d in values:
                if f"{d['node']}-{d['cycle']}" == node:
                    new = d

            for key in new.keys():
                if key == "node_type":
                    old_val = old["type"].strip(to_strip)
                elif key == "parent_ds":
                    old_val = old["parent_station"].strip(to_strip)
                else:
                    old_val = str(old[key.strip("_")]).strip(to_strip)
                new_val = str(new[key]).strip(to_strip)
                if key == "active_":
                    if float(old_val) == 1 and new_val != "True":
                        out[node] += 1
                    elif float(old_val) == 0 and new_val != "False":
                        out[node] += 1
                elif old_val != new_val:
                    out[node] += 1
        return out

    def push_changes(self, vals, changes) -> None:
        db_path = DatabaseHandler().path
        now = datetime.now().isoformat()
        tuples = []
        for d in vals:
            tuples.append(
                (
                    d["parent_ds"],
                    d["node_type"],
                    d["region"],
                    d["country"],
                    d["plan_sequencing_time"],
                    d["flex"],
                    d["cluster_comments_"],
                    d["autoassign"],
                    d["upload_scc"],
                    d["comments_"],
                    now,
                    os.getlogin(),
                    int(d["active_"]),
                    d["node"],
                    d["cycle"],
                )
            )

        tuples_audit = []
        for key in changes:
            change_count = changes[key]
            tuples_audit.append((key, os.getlogin(), now, change_count))

        with sqlite3.connect(db_path) as db:
            cursor = db.cursor()
            cursor.executemany(
                """
                UPDATE checklists
                SET parent_station = ?,
                    type = ?,
                    region = ?,
                    country = ?,
                    plan_sequencing_time = ?,
                    flex = ?,
                    cluster_comments = ?,
                    autoassign = ?,
                    upload_scc = ?,
                    comments = ?,
                    last_audit = ?,
                    audited_by = ?,
                    active = ?
                WHERE node = ? AND
                    cycle = ?;
                """,
                tuples,
            )
            cursor.executemany(
                """
            INSERT INTO audit_history (
                              checklist,
                              audited_by,
                              audited_on,
                              changes
                          )
                          VALUES (
                              ?,
                              ?,
                              ?,
                              ?
                          );
            """,
                tuples_audit,
            )

    def verify_values(self, vals: dict) -> list:
        empty_fields = []
        for key, value in vals.items():
            if key.endswith("_"):
                continue
            else:
                if isinstance(value, str):
                    if not value.strip("\t\n ;.,\\/"):
                        empty_fields.append(key)
                else:
                    if not value:
                        empty_fields.append(key)
        return empty_fields

    def main_loop(self):
        while True:
            ev, vals = self.window.read()

            if ev in (sg.WIN_CLOSED, "Cancel"):
                raise AbortedException
            elif ev == "Submit":
                empty_fields = self.verify_values(vals)
                if empty_fields:
                    s = "\n".join(empty_fields)
                    sg.PopupError(f"Empty fields detected:\n{s}")
                    continue
                formatted_vals = self.output_to_json(vals)
                changes = self.count_changes(formatted_vals)
                self.push_changes(formatted_vals, changes)
                sg.Popup("Success!", background_color="green")
                break
        self.window.close()


def main_script_checklist(node):
    try:
        ChecklistWindowMultiple(node).main_loop()
        
    except AbortedException:
        pass
    except Exception:
        logger.exception("Uncaught exception happened.")
        exit(-1)
