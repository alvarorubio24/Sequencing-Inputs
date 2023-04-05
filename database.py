import datetime
import os
import sqlite3
from pathlib import Path
from time import sleep
from typing import Iterable, Union


class DatabaseHandler:
    def __init__(self) -> None:
        self.folder = Path(
            r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\AM Shift\1. Station Checklists\1. Checklist Manager\New Checklist\Toolbox Configuration Files"
        )
        self.path = self.folder / "database.sqlite3"

    def add_checklist(
        self,
        node: str,
        parent_station: str,
        node_type: str,
        region: str,
        country: str,
        cycle: str,
        plan_sequencing_time: str,
        flex: str,
        cluster_comments: str,
        autoassign: str,
        upload_scc: str,
        comments: str,
        active: int,
    ) -> bool:
        success = False
        with sqlite3.connect(self.path) as conn:
            cursor = conn.cursor()
            retries = 0
            while retries <= 5:
                try:
                    now = datetime.datetime.now().isoformat()
                    cursor.execute(
                        """INSERT INTO checklists(
                            node,
                            parent_station,
                            type,
                            region,
                            country,
                            cycle,
                            plan_sequencing_time,
                            flex,
                            cluster_comments,
                            autoassign,
                            upload_scc,
                            comments,
                            last_audit,
                            audited_by,
                            active
                        ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (
                            node,
                            parent_station,
                            node_type,
                            region,
                            country,
                            cycle,
                            plan_sequencing_time,
                            flex,
                            cluster_comments,
                            autoassign,
                            upload_scc,
                            comments,
                            now,
                            os.getlogin(),
                            active,
                        ),
                    )
                    success = True
                    break
                except sqlite3.OperationalError as e:
                    sleep(1.5)
                    retries += 1
                    print(f"Retrying {retries}/5")
                    if retries > 5:
                        raise e

            cursor.close()
        return success

    def update_checklist(
        self,
        node: str,
        parent_station: str,
        node_type: str,
        region: str,
        country: str,
        cycle: str,
        plan_sequencing_time: str,
        flex: str,
        cluster_comments: str,
        autoassign: str,
        upload_scc: str,
        comments: str,
        active: int,
        changes: int = None,
    ) -> bool:
        success = False
        with sqlite3.connect(self.path) as conn:
            cursor = conn.cursor()
            retries = 0
            while retries <= 5:
                try:
                    now = datetime.datetime.now().isoformat()
                    cursor.execute(
                        f"""UPDATE checklists set 
                        parent_station=?,
                        type=?,
                        region=?,
                        country=?,
                        plan_sequencing_time=?,
                        flex=?,
                        cluster_comments=?,
                        autoassign=?,
                        upload_scc=?,
                        comments=?,
                        last_audit=?,
                        audited_by=?,
                        active=? WHERE node = "{node}" AND cycle = "{cycle}";""",
                        (
                            parent_station,
                            node_type,
                            region,
                            country,
                            plan_sequencing_time,
                            flex,
                            cluster_comments,
                            autoassign,
                            upload_scc,
                            comments,
                            now,
                            os.getlogin(),
                            active,
                        ),
                    )
                    cursor.execute(
                        "INSERT INTO audit_history (checklist, audited_by, audited_on, changes) VALUES (?,?,?,?)",
                        (f"{node}-{cycle}", os.getlogin(), now, changes),
                    )
                    success = True
                    break
                except sqlite3.OperationalError:
                    sleep(1.5)
                    retries += 1
                    print(f"Retrying {retries}/5")
        return success

    def add_email(self, station: str, email: str) -> bool:
        success = False
        with sqlite3.connect(self.path) as conn:
            cursor = conn.cursor()
            retries = 0
            while retries <= 5:
                try:
                    cursor.execute(
                        """INSERT INTO emails(station, email) VALUES(?, ?)""",
                        (station, email),
                    )
                    success = True
                    break
                except sqlite3.OperationalError as e:
                    print(e)
                    sleep(1.5)
                    retries += 1
                    print(f"Retrying {retries}/5")
            cursor.close()
        return success

    def update_email(self, station: str, mail: str) -> bool:
        success = False
        with sqlite3.connect(self.path) as conn:
            cursor = conn.cursor()
            retries = 0
            while retries < 5:
                try:
                    cursor.execute(
                        f'UPDATE emails SET email="{mail}" WHERE station = "{station}"'
                    )
                    success = True
                    break
                except sqlite3.OperationalError:
                    sleep(1.5)
                    retries += 1
                    print(f"Retrying {retries}/5")
            cursor.close()
        return success

    def _get_records(self, cursor: sqlite3.Cursor, query: str) -> list[dict]:
        headers = [val[0] for val in cursor.execute(query).description]
        out = []
        for row in cursor.execute(query).fetchall():
            out.append({headers[i]: row[i] for i in range(len(headers))})
        return out

    def get_station_checklists(
        self, stations: Union[Iterable, None] = None, active_only: bool = True
    ) -> list[dict]:
        station_string = '("'
        try:
            station_string += '", "'.join(stations)
        except TypeError:
            pass
        station_string += '")'
        query = "SELECT * FROM checklists"
        if active_only:
            query += " WHERE active = 1"
        if stations:
            query = (
                f"""SELECT * FROM checklists WHERE parent_station IN {station_string}"""
            )
            if active_only:
                query += " AND active = 1"
        with sqlite3.connect(self.path) as conn:
            cursor = conn.cursor()
            output = self._get_records(cursor, query)

        return output

    def get_email(self, station: str) -> str:
        with sqlite3.connect(self.path) as conn:
            cursor = conn.cursor()
            query = f'SELECT * FROM emails WHERE station = "{station}"'
            results = self._get_records(cursor, query)
        try:
            return results[0]["email"]
        except (KeyError, IndexError):
            return ""

    def check_if_checklist_exists(self, node: str, cycle: str) -> bool:
        with sqlite3.connect(self.path) as conn:
            cursor = conn.cursor()
            results = cursor.execute(
                f'SELECT * FROM checklists WHERE key = "{node}-{cycle}"'
            ).fetchall()

        return len(results) > 0

    def get_single_checklist(self, node: str, cycle: str) -> dict:
        with sqlite3.connect(self.path) as conn:
            cursor = conn.cursor()
            query = (
                f'SELECT * FROM checklists WHERE node = "{node}" AND cycle = "{cycle}"'
            )
            result = self._get_records(cursor, query)
        try:
            return result[0]
        except IndexError:
            return {}