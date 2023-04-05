import datetime
import os
import statistics
import time
from http.cookiejar import CookieJar
from math import inf
from pathlib import Path
from pprint import pprint

import browser_cookie3
import dill
import pandas as pd
import requests
from dateutil import tz
from loguru import logger
from PySimpleGUI import PopupAnnoying
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

logger.add(
    "logs\\utils.log",
    rotation="500 KB",
    format="{time:YYYY-MM-DD at HH:mm:ss} | {level} | {message}",
    backtrace=True,
    diagnose=True,
    enqueue=True,
)

nan = float("nan")


class Auth:
    def __init__(self) -> None:
        self.name = "jar.p"
        self.path = Path(os.getcwd()) / self.name
        self.domains = [
            ".amazon.com",
            ".amazonaws.com",
            ".amazon.co.uk",
            ".a2z.com",
            "w.amazon.com",
            "logistics.amazon.co.uk",
        ]
        self._patch_for_ssl()

    def create_auth(self, engine: str = "chrome") -> None:
        """Retrieves authentication using selenium, user needs to login manually."""
        chromedriver_path = "chromedriver.exe"

        try:
            os.system(f'rmdir /S /Q "{os.getcwd()}\\webdriver files\\"')
        except FileNotFoundError:
            pass

        try:
            os.makedirs("webdriver files\\firefox", exist_ok=True)
            options = webdriver.ChromeOptions()
            options.add_argument(f"user-data-dir={os.getcwd()}\\webdriver files\\")
            options.add_argument("--profile-directory=Default")
            options.add_argument("--ignore-certificate-errors")

            if engine == "chrome":
                driver = webdriver.Chrome(chromedriver_path, options=options)
            elif engine == "firefox":
                driver = webdriver.Firefox()
            else:
                raise ValueError(f'Engine must be one of {["chrome","firefox"]}')
        except WebDriverException:
            logger.exception("Webdriverexception")
            PopupAnnoying(
                "Unable to start chromedriver.\nPlease check your chromedriver version",
                title="Error",
                keep_on_top=True,
            )
        else:
            driver.get("https://midway-auth.amazon.com/")
            while f"Welcome {os.getlogin()}!" not in driver.page_source:
                time.sleep(5)
            while f"Welcome {os.getlogin()}!" in driver.page_source:
                break

            driver.get(
                "https://eu.route.planning.last-mile.a2z.com/route-planning/DBW1/761553f5-9fc1-4cef-8815-b974bc63f0a9/"
            )
            try:
                WebDriverWait(driver, 30).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "css-x3vv9y"))
                )
                time.sleep(5)
            except TimeoutException:
                logger.exception("Timeout error on RTW")
            driver.get(
                "https://eu.dispatch.planning.last-mile.a2z.com/dispatch-config/DBW1/761553f5-9fc1-4cef-8815-b974bc63f0a9/"
            )
            try:
                WebDriverWait(driver, 30).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "css-1xz1o3r"))
                )
                time.sleep(5)
            except TimeoutException:
                logger.exception("Timeout error on DPO")

            driver.get(
                "https://eu.dispatch.planning.last-mile.a2z.com/route-constraints/DBW1/761553f5-9fc1-4cef-8815-b974bc63f0a9/"
            )
            try:
                WebDriverWait(driver, 30).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "css-ly5121"))
                )
                time.sleep(5)
            except TimeoutException:
                logger.exception("Timeout error on RCS")

            driver.get("https://sim.amazon.com")
            try:
                WebDriverWait(driver, 30).until(
                    EC.visibility_of_element_located(
                        (By.CLASS_NAME, "editable-status-field")
                    )
                )
                time.sleep(5)
            except TimeoutException:
                logger.exception("Timeout error on SIM")
            driver.get(
                "https://logistics.amazon.co.uk/internal/scheduling/dsps?serviceAreaId=71c48e52-1ed3-4299-bb73-133a5241a473&date=2022-01-05"
            )
            try:
                WebDriverWait(driver, timeout=30).until(
                    EC.element_to_be_clickable((By.ID, "expandAllBtn"))
                )
                time.sleep(5)
            except TimeoutException:
                logger.exception("Timeout error on SUI")
            driver.get("https://logistics.amazon.co.uk/station/dashboard/search")
            time.sleep(5)
            jars = {}
            for domain in self.domains:
                try:
                    jar = browser_cookie3.chrome(
                        cookie_file=f"{os.getcwd()}\\webdriver files\\Default\\Cookies",
                        key_file=f"{os.getcwd()}\\webdriver files\\Local State",
                        domain_name=domain,
                    )
                    jars[domain] = jar
                except browser_cookie3.BrowserCookieError:
                    jar = browser_cookie3.chrome(
                        cookie_file=f"{os.getcwd()}\\webdriver files\\Default\\Network\\Cookies",
                        key_file=f"{os.getcwd()}\\webdriver files\\Local State",
                        domain_name=domain,
                    )
                    jars[domain] = jar
                except FileNotFoundError:
                    profile_dir = Path(driver.capabilities["moz:profile"])
                    jar = browser_cookie3.firefox(
                        profile_dir / "cookies.sqlite", domain
                    )
                    jars[domain] = jar
            with open(self.path, "wb") as f:
                dill.dump(jars, f)
        finally:
            driver.quit()

    def load_auth(self, domain: str) -> CookieJar:
        """Loads authentication for one of the allowed domains

        Args:
            domain (str): top-level domain beginning with "."

        Raises:
            KeyError: Incorrect domain requesed

        Returns:
            CookieJar: Jar to use alongside a request
        """
        if domain not in self.domains:
            raise KeyError(f"Argument domain must be one of: {self.domains}")

        try:
            with open(self.path, "rb") as f:
                jars = dill.load(f)
        except FileNotFoundError:
            return None

        return jars[domain]

    def test_auth(self) -> bool:
        jar = self.load_auth(".amazonaws.com")
        response = requests.get(
            "https://wqhfwnv4ee.execute-api.eu-west-1.amazonaws.com/ws/api/serviceTypes",
            cookies=jar,
        )
        return response.status_code == 200

    @staticmethod
    def _patch_for_ssl():
        """Forces to use proper certificates - they need to be in the same dir as file"""
        cacerts_directory = os.getcwd()
        internal_cert_filename = os.path.join(cacerts_directory, "cacerts.pem")
        merged_cert_filename = os.path.join(
            cacerts_directory, "internal_and_external_cacerts.pem"
        )

        os.environ["REQUESTS_CA_BUNDLE"] = merged_cert_filename
        os.environ["SSL_CERT_FILE"] = merged_cert_filename
        os.environ["SSL_CERT_DIR"] = cacerts_directory


class NodeData:
    def __init__(self, node: str, cycle: str, ofd_date: str, auth: Auth) -> None:
        self.node = node.upper()
        self.cycle = cycle.upper()
        self.ofd = ofd_date
        self.auth = auth
        self.jar = auth.load_auth(".amazonaws.com")
        self.cycle_id = self.get_cycle_id()

    def get_cycle_id(self) -> str:
        cycles_dict = {
            "CYCLE_0": "44efdf0b-d2e8-46ea-962c-51f5984d374a",
            "CYCLE_1": "761553f5-9fc1-4cef-8815-b974bc63f0a9",
            "CYCLE_2": "34814c9d-f499-4ebc-b93a-3c05315f5719",
            "RTS_2": "486ce335-3743-4218-9a24-44904f93e515",
            "AMEXPRESS": "e23adabe-065d-43a1-bd5a-b8e5d0c1a9a8",
            "RTS_1": "164a8ab2-bcff-4b56-8386-fd18e93c7ab8",
            "SAME_DAY": "fd9b8935-39e3-469e-acf1-9ee0e0f37e0a",
            "CYCLE_SD_A": "742c5ea0-450b-4a43-b475-2cfa732e1b47",
            "PROBLEM_SOLVE": "8f8a20a4-4808-4581-9f25-77b92e5863b5",
            "SD_1": "014e0f55-3657-4216-8876-98cda027cb3b",
            "CYCLE_EMER": "7678e5fc-f773-4cd8-ad42-77262f2b895b",
            "SD_3": "e05a29c9-afd4-48c7-bd9a-e9e51c7950bd",
            "CYCLE_4": "26cf767a-1faa-41a8-a5a2-fa2ce4c8bf1a",
            "AD_HOC_1": "20110ac5-3362-459d-8695-4c029ccfb83d",
            "CYCLE_5": "66e889f1-8083-407f-a48f-1f85599f70c2",
            "SD_2": "ac4010c1-7a1a-4a7a-ab04-d302497f6fa4",
            "AD_HOC_2": "547ac6d5-efd0-4af9-b7d4-6ea92998f3fe",
            "LOCKER_REDIRECT": "ac2d1ff8-cd2a-4771-b592-1edd16c75704",
            "CYCLE_SD_C": "814b4660-05f6-4c50-821f-08ff0e811aa9",
            "CYCLE_SD_B": "73a83e3c-bb06-4d8a-8912-9058bf718cb2",
            "CYCLE_3": "d91fe55d-721b-4cb1-9d9d-4f5e342d0642",
            "TEST_1": "091d990a-075d-4a74-8db2-25bdfa326105",
            "CRTN": "0fb22592-5c2b-4a84-ad41-8c158175d3ad",
            "SWA_AM": "f38ebf65-65f6-4cc2-a860-10bee673d581",
            "SWA_PM": "20101dc0-94ad-4fc9-aa7d-a8e9266aa2ac",
            "CYCLE_MOCO": "24b3a722-a61d-487c-9be6-6379f0c1623f",
        }
        try:
            return cycles_dict[self.cycle]
        except KeyError:
            req = requests.get(
                f"https://wqhfwnv4ee.execute-api.eu-west-1.amazonaws.com/ws/api/routingConfig/{self.node}",
                cookies=self.jar,
            )
            cycles = {
                c["cycleName"]: c["cycleId"] for c in req.json()["supportedCycles"]
            }
            return cycles[self.cycle]

    def get_dsp_definitions(self) -> pd.DataFrame:
        """Retrieves DSPs defined for station

        Returns:
            pd.DataFrame: Dataframe with DSP data
        """
        req = requests.get(
            "https://wqhfwnv4ee.execute-api.eu-west-1.amazonaws.com/ws/api/providers",
            params={"stationCode": self.node},
            cookies=self.jar,
        )
        return pd.DataFrame.from_dict(req.json())

    def get_service_types(self) -> pd.DataFrame:
        req = requests.get(
            "https://wqhfwnv4ee.execute-api.eu-west-1.amazonaws.com/ws/api/serviceTypes",
            cookies=self.jar,
        )
        df = pd.DataFrame.from_dict(req.json())
        return df[["serviceTypeName", "serviceTypeId"]]


class RtwData(NodeData):
    def __init__(self, node: str, cycle: str, ofd_date: str, auth: Auth) -> None:
        super().__init__(node, cycle, ofd_date, auth)
        self.get_data()

    def get_data(self) -> None:
        """Refreshes the data stored in class"""
        req = requests.get(
            f"https://wqhfwnv4ee.execute-api.eu-west-1.amazonaws.com/ws/api/routePlans/{self.node}/{self.cycle_id}",
            cookies=self.jar,
            params={"planDate": self.ofd, "latestOnly": False},
        )
        self.request_data = req.json()

    def get_raw_inputs(self) -> pd.DataFrame:
        """Retrieves inputs.

        Returns:
            pd.DataFrame: data with raw inputs
        """
        df = pd.DataFrame()

        if not self.request_data:
            self.inputs = df
            return pd.DataFrame()
        data = self.request_data
        for cluster in data:
            cluster_name = cluster["clusterCode"]
            for plan in cluster["routePlanSummaries"]:
                if plan["isActivePlan"] and plan["status"] == "Completed":
                    active_plan = plan
            try:
                cluster_volume = sum(active_plan["routePlanResults"]["volume"].values())
            except AttributeError:
                cluster_volume = active_plan["routePlanResults"]["volume"]
                if not active_plan["routePlanResults"]["volume"]:
                    cluster_volume = 0
            cluster_inputs = pd.DataFrame.from_dict(
                active_plan["routePlanResults"]["labour"]
            )
            cluster_inputs["Cluster volume"] = cluster_volume
            cluster_inputs["Cluster"] = cluster_name
            df = df.append(cluster_inputs)
        df["Station"] = self.node
        df["Cycle"] = self.cycle
        df["OFD Date"] = self.ofd

        return df

    def get_plan_info(self) -> dict:
        """Retrieves plan related metrics. Use get_data method first

        Returns:
            dict: dictionary with plan-related metrics
        """
        clusters_count = len(self.request_data)
        if clusters_count == 0:
            return {}
        requesters = []
        start_times = []
        finish_times = []
        plans = 0
        reasons = []
        if "errors" in self.request_data:
            logger.debug(self.request_data["errors"])
            return {}
        for cluster in self.request_data:
            for plan in cluster["routePlanSummaries"]:
                if plan.get("requestedBy") == "RPFIntegrationTest":
                    continue
                plans += 1
                if plan.get("requestedBy", None):
                    requesters.append(plan["requestedBy"])

                started = plan.get("createdDateTime", None)
                if started:
                    start_times.append(started)

                finished = plan.get("completedDateTime")
                if finished:
                    finish_times.append(finished)

                reason = plan.get("reason", None)
                if reason:
                    reasons.append(reasons)
        try:
            scheduler = statistics.mode(requesters)
        except statistics.StatisticsError:
            scheduler = ", ".join(tuple(set(requesters)))
        local = tz.tzwinlocal()
        out = {
            "started": datetime.datetime.fromtimestamp(
                float(f"{min(start_times)}"[:-3]), tz=local
            ),
            "plans": plans,
            "replans": plans - clusters_count,
            "requesters": tuple(set(requesters)),
            "scheduler": scheduler,
        }

        try:
            out["finished"] = datetime.datetime.fromtimestamp(
                float(f"{max(finish_times)}"[:-3]), tz=local
            )
        except Exception:
            pass

        return out

    def get_readable_inputs(self, skip_service_type_merge=False) -> pd.DataFrame:
        """Gets inputs and merges with DSP definitions and optionally with service types

        Args:
            skip_service_type_merge (bool, optional): if retrieving more than one set of inputs, it may be more efficient to skip this merge and do at the end. Defaults to False.

        Returns:
            pd.DataFrame: dataframe with inputs
        """
        inputs = self.get_raw_inputs()
        if inputs.empty:
            return pd.DataFrame()
        dsp = self.get_dsp_definitions()
        df = pd.merge(
            inputs,
            dsp[["companyId", "companyShortCode"]],
            left_on="dspId",
            right_on="companyId",
            how="left",
        )
        if not skip_service_type_merge:
            service_types = self.get_service_types()
            df = pd.merge(
                df,
                service_types[["serviceTypeId", "serviceTypeName"]],
                on="serviceTypeId",
                how="left",
            )
        for col in ["minRoutes", "maxRoutes"]:
            df[col] = df[col].apply(lambda x: float("inf") if x >= 1000000 else x)
        df["Reduce by"] = df["maxRoutes"] - df["minRoutes"]
        df.rename(
            columns={
                "serviceTypeName": "Service type",
                "companyShortCode": "DSP",
                "localDepartureTime": "Depart time",
                "shiftTimeMinutes": "Shift time",
                "maxRoutes": "Requested",
                "actualRoutes": "Planned",
                "spr": "SPR",
                "avgCubeUtilization": "Vehicle Util.",
                "avgTimeUtilization": "Time Util.",
            },
            inplace=True,
        )
        df["Cuts"] = df["Requested"] - df["Planned"]
        df["Cuts"] = df["Cuts"].apply(lambda x: 0 if x in [inf, -inf] else x)
        if skip_service_type_merge:
            return df[
                [
                    "serviceTypeId",
                    "DSP",
                    "Depart time",
                    "Shift time",
                    "Requested",
                    "Reduce by",
                    "Planned",
                    "SPR",
                    "Vehicle Util.",
                    "Time Util.",
                    "Cluster",
                    "Cluster volume",
                    "Cuts",
                    "OFD Date",
                ]
            ].copy()

        return df[
            [
                "Service type",
                "DSP",
                "Depart time",
                "Shift time",
                "Requested",
                "Reduce by",
                "Planned",
                "SPR",
                "Vehicle Util.",
                "Time Util.",
                "Cluster",
                "Cluster volume",
                "Cuts",
                "OFD Date",
            ]
        ].copy()


class DpoData(NodeData):
    def __init__(self, node: str, cycle: str, ofd_date: str, auth: Auth) -> None:
        super().__init__(node, cycle, ofd_date, auth)

    def get_data(self):
        req = requests.get(
            f"https://0vxi8gngy9.execute-api.eu-west-1.amazonaws.com/ws/api/dispatchPlan/{self.node}/{self.cycle_id}/{self.ofd}",
            cookies=self.jar,
        )
        return req.json()

    def get_unplanned_routes(self):
        js = self.get_data()
        try:
            dsp_unplanned = len(js.get("dspPlan", {}).get("unplannedRoutes", {}))
        except AttributeError:
            dsp_unplanned = 0
        flex_unplanned = 0
        flex_not_considered = 0
        if js.get("flexPlan", None):
            flex_unplanned = len(js.get("flexPlan", {}).get("unplannedRoutes", 0))
            for block in js["flexPlan"]["blocks"]:
                prescheduled = block["prescheduledCount"]
                planned = block["plannedCount"]
                if prescheduled > planned:
                    flex_not_considered += prescheduled - planned

        return {
            "dsp_unplanned": dsp_unplanned,
            "flex_unplanned": flex_unplanned,
            "flex_not_considered": flex_not_considered,
        }


class RcsData(NodeData):
    def __init__(self, node: str, cycle: str, ofd_date: str, auth: Auth) -> None:
        super().__init__(node, cycle, ofd_date, auth)

    def get_data(self) -> pd.DataFrame:
        req = requests.get(
            "https://0vxi8gngy9.execute-api.eu-west-1.amazonaws.com/ws/api/routeConstraints",
            params={
                "stationCode": self.node,
                "cycleId": self.cycle_id,
                "planDate": self.ofd,
            },
            cookies=self.jar,
        )
        js = req.json()
        dfs = []
        try:
            for cluster in js["constraintsByCluster"]:
                df = pd.DataFrame.from_records(js["constraintsByCluster"][cluster])
                df["Cluster"] = cluster
                dfs.append(df)
            return pd.concat(dfs)
        except KeyError:
            pprint(js)

    def get_inputs(self, infinity_value=float("inf")) -> pd.DataFrame:
        df = self.get_data()
        services = self.get_service_types()
        providers = self.get_dsp_definitions()
        df = pd.merge(df, services, on="serviceTypeId", how="left")
        df = pd.merge(
            df,
            providers[["companyId", "companyShortCode"]],
            left_on="providerId",
            right_on="companyId",
            how="left",
        )
        df["Reduce by"] = df["max"] - df["min"]
        df.drop(columns=["serviceTypeId", "companyId", "providerId"], inplace=True)
        df.rename(
            columns={
                "departTime": "Depart time",
                "scheduledDepartTime": "Scheduled time",
                "companyShortCode": "DSP",
                "serviceTypeName": "Service type",
                "onRoadMinutes": "Route length",
                "scheduledMinutes": "Shift length",
                "max": "Scheduled",
            },
            inplace=True,
        )
        df["Reduce by"].fillna(0, inplace=True)
        df["Scheduled"].fillna(infinity_value, inplace=True)
        column_order = [
            "Service type",
            "DSP",
            "Shift length",
            "Route length",
            "Scheduled time",
            "Depart time",
            "Scheduled",
            "Reduce by",
            "Cluster",
        ]
        other_cols = [col for col in df.columns if col not in column_order]
        return df[[*column_order, *other_cols]]