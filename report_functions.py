import datetime
import os
from typing import Iterable

import pandas as pd
import win32com.client
from bs4 import BeautifulSoup
from dateutil import tz
from loguru import logger
from numpy import inf, nan

import amzl_requests
from cache_data import cache_dir

logger.add(
    "logs\\test.log",
    rotation="10 MB",
    format="{time:YYYY-MM-DD at HH:mm:ss} | {level} | {message}",
    backtrace=True,
    diagnose=True,
    enqueue=True,
)


def get_hyperzone_clusters(cluster_names: Iterable) -> list:
    """Checks which clusters are Hyperzone (HZ). Needed to separate calculations from other clusters.
    Mostly useless, as the HZ cluster were deprecated a while ago.

    Args:
        cluster_names (Iterable): cluster names to verify

    Returns:
        list: list of hyperzone clusters
    """
    definitions = pd.read_csv(
        r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\AM Shift\1. Station Checklists\1. Checklist Manager\New Checklist\Toolbox Configuration Files\hyperzone_clusters.csv"
    )

    hyperzone_clusters = []
    # Excluded clusters
    excluded = definitions.loc[definitions["match_type"] == "exclude"]

    if not excluded.empty:
        excluded = excluded["match_type"].values
    else:
        excluded = []

    # Fuzzy matches
    fuzzy = definitions.loc[definitions["match_type"] == "fuzzy"]
    if not fuzzy.empty:
        for c in fuzzy["cluster_name"]:
            for cluster in cluster_names:
                if c in cluster.lower():
                    if cluster not in excluded:
                        hyperzone_clusters.append(cluster)

    # Exact matches
    exact = definitions.loc[definitions["match_type"] == "exact"]
    if not exact.empty:
        for c in exact["cluster_name"]:
            if c in cluster_names:
                if cluster not in excluded:
                    hyperzone_clusters.append(c)

    return hyperzone_clusters


def summarise_inputs(inputs: pd.DataFrame) -> dict:
    """Summarises inputs in 4 metrics - cuts, infinity routes, minutes used (for each service type) and creates a pivot table with detailed summary

    Args:
        inputs (pd.DataFrame): raw inputs from

    Returns:
        dict: summary
    """

    def append_hz(vals):
        """Appends " [HZ]" to separate normal service type from hyperzone service type.
        Args:
            vals (tuple): Service type and cluster name

        Returns:
            str: service type
        """
        service_type, cluster = vals
        if cluster in hyperzone_clusters:
            service_type += r" [HZ]"
        return service_type

    df = inputs
    df["Volume"] = df["SPR"] * df["Planned"]
    hyperzone_clusters = get_hyperzone_clusters(cluster_names=df["Cluster"].unique())
    df["Service type"] = df[["Service type", "Cluster"]].apply(append_hz, axis=1)

    # Calculating actual volume for each line (since SPR figure is rounded, it would create inconsistencies)
    cluster_volumes = df.pivot_table(index="Cluster", values="Volume", aggfunc="sum")[
        "Volume"
    ].to_dict()
    df["_cluster_volume_fake"] = df["Cluster"].apply(lambda x: cluster_volumes[x])
    df["Volume share"] = df["Volume"] / df["_cluster_volume_fake"]
    df["Volume"] = df["Cluster volume"] * df["Volume share"]

    df.replace({"∞": inf}, inplace=True)
    pivot = df.pivot_table(
        index="Service type", values=["Volume", "Planned"], aggfunc="sum"
    )
    # IMPORTANT: we cannot use average of SPR column here, as it would not take into account the amount of routes generated with given SPR.
    # Therefore we have to sum up volume and routes and calculate it that way.
    pivot["Volume"] = pivot["Volume"].round(0)
    total_volume = pivot["Volume"].sum()
    pivot.loc["Total"] = pivot.sum(numeric_only=True, axis=0)
    pivot["SPR"] = pivot["Volume"] / pivot["Planned"]
    pivot["Share %"] = pivot["Volume"] / total_volume
    pivot["Volume"] = pivot["Volume"].apply(lambda x: f"{x:.0f}")
    pivot["Planned"] = pivot["Planned"].apply(lambda x: f"{x:.0f}")
    pivot.rename(columns={"Planned": "Routes"}, inplace=True)
    pivot.reset_index(inplace=True, drop=False)
    pivot = pivot[["Service type", "Routes", "SPR", "Volume", "Share %"]]

    # Creating a string which lists minutes used for each service type.
    minutes = []
    for service in df.sort_values("Service type", ascending=True, inplace=False)[
        "Service type"
    ].unique():
        minutes_used = (
            df.loc[df["Service type"] == service]["Shift time"].unique().tolist()
        )
        minutes_used = [str(minutes) for minutes in minutes_used]
        minutes.append(f"{service}: {'/'.join(minutes_used)}")
    minutes = "\n".join(minutes)
    minutes = minutes.strip(" \n.,")

    df["Infinity Routes"] = df["Planned"]
    df.loc[df["Requested"] != inf, "Infinity Routes"] = 0
    infinity_routes = df["Infinity Routes"].sum()
    cuts = df["Cuts"].sum()

    return {
        "minutes": minutes,
        "cuts": int(cuts),
        "infinity": int(infinity_routes),
        "package_count": pivot,
    }


def summarise_plans(plans: dict, country: str) -> dict:
    """Summarising plan details (SLAs, replans etc)

    Args:
        plans (dict): plan data
        country (str): Country of station, needed for local time calculation

    Returns:
        dict: summary
    """

    # Calculating local time (adjusts for DST)
    # Not pretty, but other options stopeed working for some people
    # This implementation works for everyone
    local = tz.tzlocal()
    dst = datetime.datetime.now(local).dst().seconds > 0
    local_offset = datetime.datetime.now(local).utcoffset().seconds
    dst_adjust = {False: 3600, True: 7200}
    dst_no_adjust = {False: 0, True: 3600}
    timezones = {
        "UK": dst_no_adjust,
        "IR": dst_no_adjust,
        "IE": dst_no_adjust,
        "AT": dst_adjust,
        "ES": dst_adjust,
        "IT": dst_adjust,
        "FR": dst_adjust,
        "DE": dst_adjust,
        "NL": dst_adjust,
        "BE": dst_adjust,
    }

    offset = timezones.get(country, dst_adjust).get(dst)
    replans = plans.get("replans", None)
    try:
        start_time = plans.get("started", None)
        start_time = start_time + datetime.timedelta(seconds=offset - local_offset)
        start_time = start_time.strftime("%H:%M")
    except Exception:
        start_time = None

    try:
        finish_time = plans.get("finished", None)
        finish_time = finish_time + datetime.timedelta(seconds=offset - local_offset)
        finish_time = finish_time.strftime("%H:%M")
    except Exception:
        finish_time = None

    if plans.get("scheduler", None):
        scheduler = plans["scheduler"]
    else:
        scheduler = ", ".join(plans.get("requesters", [" "]))

    logger.info(f"Start time: {start_time}")
    logger.info(f"Finish time: {finish_time}")
    return {
        "replans": replans,
        "start_time": start_time,
        "finish_time": finish_time,
        "scheduler": scheduler,
    }


def create_html(fields: dict) -> str:
    """Creates HTML for sequencing report email body

    Args:
        fields (dict): data to include

    Returns:
        str: HTML email body
    """
    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style type="text/css">
        table {
    font-family: Verdana, Geneva, sans-serif;
    border: 2px solid #000000;
    background-color: #EEEEEE;
    text-align: left;
    border-collapse: collapse;
    }
    table td, table th {
    border: 1px solid #000000;
    padding: 3px 2px;
    word-break:normal;
    }
    table tbody td {
    font-size: 10px;
    color: #333333;
    }
    table tr:nth-child(even) {
    background: #D0E4F5;
    }
    table thead {
    background: #0B6FA4;
    border-bottom: 1px solid #000000;
    }
    table thead th {
    font-size: 12px;
    font-weight: bold;
    color: #FFFFFF;
    text-align: center;
    border-left: 1px solid #000000;
    height: 80px;
    }
    table thead th:first-child {
    border-left: none;
    }

    table tfoot {
    font-size: 10px;
    font-weight: bold;
    color: #333333;
    background: #D0E4F5;
    border-top: 3px solid #444444;
    }
    table tfoot td {
    font-size: 10px;
    }

    html {
      font-size: 62.5%; }
    body {
      font-size: 1.5em; /* currently ems cause chrome bug misinterpreting rems on body element */
      line-height: 1.6;
      font-weight: 400;
      font-family: "Raleway", "HelveticaNeue", "Helvetica Neue", Helvetica, Arial, sans-serif;
      color: #222; }


    /* Typography
    –––––––––––––––––––––––––––––––––––––––––––––––––– */
    h1, h2, h3, h4, h5, h6 {
      margin-top: 0;
      margin-bottom: 2rem;
      font-weight: 300; }
    h1 { font-size: 4.0rem; line-height: 1.2;  letter-spacing: -.1rem;}
    h2 { font-size: 3.6rem; line-height: 1.25; letter-spacing: -.1rem; }
    h3 { font-size: 3.0rem; line-height: 1.3;  letter-spacing: -.1rem; }
    h4 { font-size: 2.4rem; line-height: 1.35; letter-spacing: -.08rem; }
    h5 { font-size: 1.8rem; line-height: 1.5;  letter-spacing: -.05rem; }
    h6 { font-size: 1.5rem; line-height: 1.6;  letter-spacing: 0; }
    </style>
    </head>
    <body>
    """
    message_header = f"Hello {fields['Station']} team,<br><p>Please find below the sequencing report for today.</p>"
    message_footer = '<p><span style="color: #999999;"><br>Email sent automatically. If there are any technical issues with this report please get in contact with <em>@kcmicha</em></span></p></body></html>'
    tables = ["<h4><strong>Summary:</strong></h4>"]
    tables.append(
        fields["Summary"]
        .rename(
            columns={
                "Planned induct finish": "Planned sequencing time",
                "Induct finish time": "Request sequencing time",
                "Induct time difference": "Request time difference",
            }
        )
        .to_html(
            classes="summary",
            escape=False,
            index=False,
            columns=[
                "Wave",
                "Planned sequencing time",
                "Request sequencing time",
                "Request time difference",
                "Sequence start time",
                "Sequence finish time",
                "Routes assigned at",
                "DPO files sent at",
                "Sequencing time",
                "Precuts",
                "Cuts",
                "Infinity routes",
                "Minutes",
                "Replans",
                "Comment",
            ],
            na_rep="",
        )
    )

    for key in fields:
        if key != "Station" and key != "Summary":
            tables.append(f"<br><br><h4><strong>{key}:</strong></h4>")
            if fields[key]["forecast_volume"] is not nan:
                tables.append(
                    f"Forecast volume: {fields[key]['forecast_volume']}<br><br>"
                )
            tables.append(
                fields[key]["df"].to_html(
                    escape=False, classes="cycle", index=False, na_rep=" "
                )
            )
            # with open("test.html", "w") as f:
            #     f.write(
            #         fields[key]["df"].to_html(
            #             escape=False, classes="cycle", index=False, na_rep=" "
            #         )
            #     )

    return html + message_header + "".join(tables) + message_footer


def replace_newline(val: str):
    try:
        if "\n" in str(val):
            return str(val).replace("\n", "<br>")
        return val
    except Exception:
        return val


def compare_datetimes(values) -> datetime.timedelta:
    """Calculates SLA, needed because sequencing report does not store date information, only time

    Args:
        values (tuple): start and finish time

    Returns:
        datetime.timedelta: calculated SLA
    """
    start, finish = values
    try:
        if (
            start > finish
            and start.hour > finish.hour
            and start.hour > 20
            and finish.hour < 20
        ):
            return finish - (start - datetime.timedelta(days=1))
        elif (
            start < finish
            and start.hour < finish.hour
            and start.hour < 20
            and finish.hour > 20
        ):
            return finish - datetime.timedelta(days=1) - start
        return finish - start
    except Exception:
        return nan


def convert_to_time(val):
    try:
        return val.strftime("%H:%M")
    except Exception:
        return nan


def gather_tables(files_path: str, station: str, ofd_date: str = None):

    # Checks tables which have sequencing summary for each station and cycle
    # May have more than one station due to XPTs
    tbls = {
        file[:-4]: f"{files_path}\\{file}"
        for file in os.listdir(files_path)
        if file.endswith(".csv")
        and not file.startswith("summary")
        and "data" not in file
    }

    tbls["Station"] = station
    tbls["Summary"] = f"{files_path}\\summary.csv"

    for key in tbls:
        if (
            key != "Summary" and key != "Station"
        ):  # Cycle summaries (data from sequencing report)
            ds = key[: key.find("-")]
            cycle = key[key.find("-") + 1 :]
            df = pd.read_csv(tbls[key])
            df["SPR"] = df["SPR"].apply(lambda x: f"{x:.0f}")
            df["Share %"] = df["Share %"].apply(lambda x: f"{x:.2%}")
            df.rename(columns={"Planned": "Routes"}, inplace=True)

            forecast = get_forecast_data(
                station=ds,
                cycle=cycle,
                siphon_data=f"{files_path}\\siphon_data.csv",
                date=ofd_date,
            )
            forecast_volume = get_volume(
                ds=ds, cycle=cycle, data=f"{files_path}\\volume_data.csv"
            )
            flex = get_flex_data(
                flex_data=f"{files_path}\\flex_data.csv",
                ds=ds,
                cycle=cycle,
                files_path=files_path,
            )

            if forecast.empty:
                logger.warning("FORECAST DATA IS EMPTY")
                tbls[key] = {
                    "df": df.replace({"nan": 0, "": 0, "0": 0}).fillna(0),
                    "forecast_volume": forecast_volume,
                }
            else:
                merge = pd.merge(
                    df,
                    forecast.rename(columns={"service_type": "Service type"}),
                    how="outer",
                    on="Service type",
                )
                merge.replace({"Total": "zzzzzz"}, inplace=True)
                merge.sort_values(by="Service type", inplace=True)
                merge.replace({"zzzzzz": "Total"}, inplace=True)
                merge = merge[
                    [
                        "Service type",
                        "Routes",
                        "SPR",
                        "Volume",
                        "Share %",
                        "Forecast Routes",
                        "Forecast SPR",
                    ]
                ]
                merge.fillna(0, inplace=True)
                for service_type in merge["Service type"].values:

                    # Flex forecast data is not with other service types, therefore need to manually merge them
                    if (
                        "AmFlex" in str(service_type)
                        and "Large Van" not in str(service_type)
                        and "Large Vehicle" not in str(service_type)
                    ):
                        merge.loc[
                            merge["Service type"] == service_type, "Forecast SPR"
                        ] = flex["spr"]
                        merge.loc[
                            merge["Service type"] == service_type, "Forecast Routes"
                        ] = flex["routes"]

                    # AmFlex Large Van and vehicle are separate from the normal flex forecast
                    # Therefore we are filling Forecast routes with actual routes (as we do not generate extras)
                    elif str(service_type) in [
                        "AmFlex Large Van",
                        "AmFlex Large Vehicle",
                    ]:
                        merge.loc[
                            merge["Service type"] == service_type, "Forecast Routes"
                        ] = flex["van_routes"]
                        merge.loc[
                            merge["Service type"] == service_type, "Forecast SPR"
                        ] = flex["van_spr"]
                        merge.loc[
                            merge["Service type"] == service_type, "Forecast Capacity"
                        ] = (flex["van_routes"] * flex["van_spr"])
                    # Hyperzone is listed differently as well, need to merge manually
                    elif "[HZ]" in str(service_type):
                        try:
                            hz_routes = merge.loc[
                                merge["Service type"] == service_type, "Forecast Routes"
                            ].values[0]
                            hz_volume = get_volume(
                                ds=ds,
                                cycle=cycle,
                                data=f"{files_path}\\volume_data.csv",
                                output="hz",
                            )
                            merge.loc[
                                merge["Service type"] == service_type, "Forecast SPR"
                            ] = (hz_volume / hz_routes)
                        except ZeroDivisionError:
                            pass
                        except Exception:
                            logger.exception("Uncaught exception with HZ calculation")

                merge.loc[merge["Service type"] == "Total", "Forecast SPR"] = (
                    merge["Forecast SPR"] * merge["Forecast Routes"]
                ).sum() / merge["Forecast Routes"].sum()
                merge.loc[merge["Service type"] == "Total", "Forecast Routes"] = merge[
                    "Forecast Routes"
                ].sum()
                merge["Forecast Capacity"] = (
                    merge["Forecast SPR"] * merge["Forecast Routes"]
                )

                for service_type in merge["Service type"].values:
                    if (
                        "AmFlex" in str(service_type)
                        and "Large Van" not in str(service_type)
                        and "Large Vehicle" not in str(service_type)
                    ):
                        merge.loc[
                            merge["Service type"] == service_type, "Forecast Capacity"
                        ] = flex["volume"]

                merge.loc[
                    merge["Service type"] == "Total", "Forecast Capacity"
                ] = merge.loc[
                    merge["Service type"] != "Total", "Forecast Capacity"
                ].sum()
                merge.fillna(0, inplace=True)
                merge.replace({"nan": 0, "nan%": 0}, inplace=True)
                merge["Forecast Share %"] = (
                    merge["Forecast Capacity"]
                    / merge.loc[
                        merge["Service type"] != "Total", "Forecast Capacity"
                    ].sum()
                )
                merge["Forecast SPR"] = merge["Forecast SPR"].apply(
                    lambda x: f"{x:.0f}"
                )
                merge["Forecast Routes"] = merge["Forecast Routes"].apply(
                    lambda x: f"{x:.0f}"
                )
                merge["Forecast Share %"] = merge["Forecast Share %"].apply(
                    lambda x: f"{x:.2%}"
                )
                merge["Forecast Capacity"] = merge["Forecast Capacity"].apply(
                    lambda x: f"{x:.0f}"
                )
                merge.fillna("0", inplace=True)
                merge.replace({"nan": "0", "nan%": "0"}, inplace=True)

                tbls[key] = {
                    "df": merge.fillna("0"),
                    "forecast_volume": forecast_volume,
                }

        if key == "Summary":  # summary table (the one at the top)
            df = pd.read_csv(tbls[key])
            df = df.loc[df["not_run"] == False].copy()
            df.replace("nan", nan, inplace=True)
            df["minutes"] = df["minutes"].apply(replace_newline)
            df["comment"] = df["comment"].apply(replace_newline)
            df["forced_induct_comment"] = df["forced_induct_comment"].apply(
                replace_newline
            )
            df["pift"] = df["pift"].apply(parse_time)
            df["aift"] = df["aift"].apply(parse_time)
            df["Induct time difference"] = df[["pift", "aift"]].apply(
                compare_datetimes, axis=1
            )
            df["Induct time difference"] = df["Induct time difference"].apply(
                lambda x: f"{x.total_seconds()//60:.0f}"
                if type(x) is not float
                else nan
            )
            df["Induct time difference"] = df["Induct time difference"].apply(
                lambda x: f"{x} min" if x != "nan" else ""
            )
            df["pift"] = df["pift"].apply(convert_to_time)
            df["aift"] = df["aift"].apply(convert_to_time)
            df.fillna(" ", inplace=True)
            df.replace(
                {"nan:nan min": " ", "nan:nan": " ", "nan": " ", "nan min": " "},
                inplace=True,
            )
            df.rename(
                columns={
                    "Unnamed: 0": "Wave",
                    "pift": "Planned induct finish",
                    "aift": "Induct finish time",
                    "sequence_start": "Sequence start time",
                    "sequence_finish": "Sequence finish time",
                    "assigned": "Routes assigned at",
                    "dpo": "DPO files sent at",
                    "precuts": "Precuts",
                    "cuts": "Cuts",
                    "replans": "Replans",
                    "minutes": "Minutes",
                    "infinity": "Infinity routes",
                    "comment": "Comment",
                    "SLA": "Sequencing time",
                    "forced_induct": "Forced induct",
                    "forced_induct_comment": "Forced induct reason",
                },
                inplace=True,
            )
            tbls["Summary"] = df
    return tbls


def parse_time(time: str) -> datetime.datetime:
    """Converting input time to datetime format.

    Args:
        time (str): Accepts different timelike formats, eg. 0700/700/07:00,7:00 for 7am

    Returns:
        datetime.datetime: datetime with today's date
    """
    try:
        today = datetime.datetime.now()
        month = today.month
        year = today.year
        day = today.day
        if ":" in time:
            hour = int(time.split(":")[0])
            minute = int(time.split(":")[1])
        else:
            if len(time) == 3:
                time += "0" + time
                hour = int(time[:2])
                minute = int(time[2:])

            elif len(time) == 4:
                hour = int(time[:2])
                minute = int(time[2:])

            else:
                return nan

        if hour == 24:
            hour = 0

        return datetime.datetime(year, month, day, hour, minute, 0)
    except Exception:
        return nan


def fill_dpo_time(vals):
    assigned, dpo = vals
    if not dpo:
        return assigned
    return dpo


def parse_package_count(csv_path: str) -> pd.DataFrame:
    """Parsing package count CSV from autoassign page. Only used in case automatic retrieval from RTW fails (which should not happen at all)

    Args:
        csv_path (str): path to the csv file
    Returns:
        pd.DataFrame: summary of the actual output
    """
    df = pd.read_csv(csv_path)
    pivot = df.pivot_table(
        index="SERVICE TYPE",
        values=["Package Count", "ROUTE"],
        aggfunc={"Package Count": "sum", "ROUTE": lambda x: x.value_counts().count()},
    )
    total_volume = pivot["Package Count"].sum()
    pivot.loc["Total"] = pivot.sum(numeric_only=True, axis=0)
    pivot["SPR"] = pivot["Package Count"] / pivot["ROUTE"]
    pivot["Share %"] = pivot["Package Count"] / total_volume
    pivot.reset_index(inplace=True)
    pivot.rename(
        columns={
            "ROUTE": "Routes",
            "Package Count": "Volume",
            "SERVICE TYPE": "Service type",
        },
        inplace=True,
    )

    return pivot[["Service type", "Routes", "Volume", "SPR", "Share %"]]


def parse_report_inputs(df: pd.DataFrame):
    for col in ["aift", "sequence_start", "sequence_finish", "assigned", "dpo"]:
        df[col] = df[col].apply(parse_time)
    # df["SLA"] = df["assigned"] - df["aift"]
    df["SLA"] = df[["aift", "assigned"]].apply(calculate_sla, axis=1)
    for col in ["aift", "sequence_start", "sequence_finish", "assigned", "dpo"]:
        try:
            df[col] = df[col].apply(lambda x: f"{x.hour:02}:{x.minute:02}")
        except Exception:
            df[col] = nan
    df["SLA"] = df["SLA"].apply(lambda x: f"{x//60:.0f} min")
    return df


def create_email(
    country: str, email_body: str, ds: str, email: str, date: datetime.date = None
):
    """Creates email in outlook

    Args:
        country (str): country of the station
        email_body (str): HTML to use as a body
        ds (str): station name
        email (_type_): email address to send to
        date (_type_, optional): OFD date of sequencing report. Defaults to datetime.date.today().
    """
    if not date:
        date = datetime.date.today()

    country_emails = {
        "UK": "eu-co-sequencing-uk@amazon.com",
        "AT": "eu-co-sequencing-meu@amazon.com",
        "ES": "eu-co-sequencing-es@amazon.com",
        "IT": "eu-co-sequencing-it@amazon.com",
        "FR": "eu-co-sequencing-fr@amazon.com",
        "DE": "eu-co-sequencing-meu@amazon.com",
        "IR": "eu-co-sequencing-uk@amazon.com",
        "NL": "eu-co-sequencing-meu@amazon.com",
        "BE": "eu-co-sequencing-fr@amazon.com",
    }
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = (
        f"{country.upper()} - {ds.upper()} - Sequencing Report - OFD: {date}"
    )
    newMail.BodyFormat = 2  # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
    newMail.HTMLBody = email_body
    newMail.To = email
    newMail.CC = country_emails.get(country.upper(), "")
    newMail.display(True)
    # newMail.send()


def calculate_sla(vals, flag=None):
    start, finish = vals
    if flag:
        seconds = (finish - start).total_seconds()
        if start > finish:
            seconds = (finish - (start - datetime.timedelta(days=1))).total_seconds()
        return f"{(seconds//60):.0f} min"
    try:

        if start > finish:
            return (finish - (start - datetime.timedelta(days=1))).total_seconds()
        return (finish - start).total_seconds()

    except Exception:
        return nan


def replace_service_types(vals, refresher: bool = True) -> str:
    """logic to replace forecast service type to match actual service type

    Args:
        vals (tuple): service type and scheduling service type
        refresher (bool, optional): refresher needs to escape markdown formatting. Defaults to True.

    Returns:
        str: Service type as named in RTW
    """
    service_type, scheduling_service_type = vals
    service_type = str(service_type)
    scheduling_service_type = str(scheduling_service_type)
    if (
        "Nursery Route Level 1" in service_type
        and "Low Emission" not in service_type
        and "Walker" not in service_type
        and "Cargo" not in service_type
    ):
        return "Nursery Route Level 1"
    elif (
        "Nursery Route Level 2" in service_type
        and "Low Emission" not in service_type
        and "Walker" not in service_type
        and "Cargo" not in service_type
    ):
        return "Nursery Route Level 2"
    elif (
        "Nursery Route Level 3" in service_type
        and "Low Emission" not in service_type
        and "Walker" not in service_type
        and "Cargo" not in service_type
    ):
        return "Nursery Route Level 3"
    elif "DSP2.0 SP Medium - Late Dispatch" in service_type:
        return "Box Truck Parcel (Medium)"
    elif "DSP2.0" in service_type:
        return scheduling_service_type
    elif "Retrain" in service_type:
        return scheduling_service_type
    elif "Ride Along" in service_type:
        return scheduling_service_type
    elif "Standard Parcel - Low Emission Vehicle with Helper" in service_type:
        return scheduling_service_type
    elif "Standard Parcel - Low Emission Vehicle" in service_type:
        return scheduling_service_type
    elif ("hyperzone" in service_type.lower()) and refresher:
        return scheduling_service_type + r" \[HZ\]"
    elif "hyperzone" in service_type.lower() and not refresher:
        return scheduling_service_type + " [HZ]"
    elif "ORDT Extra large cargo van" in service_type and refresher:
        return r"ORDT Extra large cargo van \[HZ\]"
    elif "ORDT Extra large cargo van" in service_type and not refresher:
        return r"ORDT Extra large cargo van [HZ]"
    elif "Nursery Route Level 1" in service_type and "Low Emission" in str(
        service_type
    ):
        return "Nursery Route Level 1 - Low Emissions Vehicle"
    elif "Nursery Route Level 2" in service_type and "Low Emission" in str(
        service_type
    ):
        return "Nursery Route Level 2 - Low Emissions Vehicle"
    elif "Nursery Route Level 3" in service_type and "Low Emission" in str(
        service_type
    ):
        return "Nursery Route Level 3 - Low Emissions Vehicle"
    elif "ORDT Long route" in service_type:
        return scheduling_service_type
    elif (
        "Cargo" in service_type
        and "Nursery Route Level 2" in service_type
        and "Large" in service_type
    ):
        return "Cargo Electric Bicycle Large - Nursery Route Level 2"
    elif (
        "Cargo" in service_type
        and "Nursery Route Level 1" in service_type
        and "Large" in service_type
    ):
        return "Cargo Electric Bicycle Large - Nursery Route Level 1"
    elif (
        "Cargo" in service_type
        and "Nursery Route Level 3" in service_type
        and "Large" in service_type
    ):
        return "Cargo Electric Bicycle Large - Nursery Route Level 3"
    elif str(service_type) in (
        "Standard Parcel - Low Emission Vehicle (Medium) - Long Range",
        "Standard Parcel - Low Emission Vehicle (Large) - Long Range",
        "Standard Parcel - Low Emission Vehicle (Small) - Long Range",
        "Standard Parcel - Low Emission Vehicle (Medium) - Medium Range",
    ):
        return service_type
    return service_type


def get_forecast_data(station: str, cycle: str, siphon_data, date: str = None):

    if date is None:
        today = datetime.date.today()
        date = today.strftime("%Y-%m-%d")
    try:
        if isinstance(siphon_data, str):
            df = pd.read_csv(
                siphon_data,
                converters={"ofd_date": str},
            )
        elif isinstance(siphon_data, pd.DataFrame):
            df = siphon_data
        else:
            print("siphon data must be either a path to csv or a dataframe")
            raise ValueError
        df["ofd_date"] = df["ofd_date"].apply(lambda x: x[:10])
        df = df.loc[
            (df["plan_type"] != "weekly-run")
            & (df["station"] == station)
            & (df["cycle"] == cycle)
            & (df["ofd_date"] == date)
        ]
        df.sort_values("snapshot_id", ascending=True, inplace=True)
        df.drop_duplicates(
            [
                "ofd_date",
                "station",
                "cycle",
                "service_type",
                "scheduling_service_type",
                "shift_length",
            ],
            keep="last",
            inplace=True,
        )
        df["service_type"] = df[["service_type", "scheduling_service_type"]].apply(
            replace_service_types, axis=1, refresher=False
        )
        df["Volume"] = df["spr"] * df["routes_output"]
        pivot = df.pivot_table(
            index="service_type",
            values=["Volume", "routes_output"],
            aggfunc={"Volume": "sum", "routes_output": "sum"},
        )
        pivot["spr"] = pivot["Volume"] / pivot["routes_output"]
        pivot = pivot[["routes_output", "spr"]]
        pivot.rename(
            columns={"routes_output": "Forecast Routes", "spr": "Forecast SPR"},
            inplace=True,
        )

        return pivot.reset_index()
    except Exception:
        return pd.DataFrame(columns=["service_type", "Forecast Routes", "Forecast SPR"])


def download_siphon_data(nodes: list, ofd_date: str = None):
    def select_latest_forecast(vals, snapshot_dict: dict) -> int:
        node, cycle, snapshot_id = vals
        return int(snapshot_dict.get((node, cycle), {}).get("snapshot_id", 0)) == int(
            snapshot_id
        )

    if ofd_date is None:
        ofd_date = datetime.datetime.today().strftime("%Y-%m-%d")

    try:
        df = pd.read_csv(
            r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\AM Shift\6. Handover Report\New Scheduling Model for all countries - Morning report\SIPHON Logs - updated\siphon_pull_both_service_types.txt",
            sep="\t",
        )
        df["ofd_date"] = df["ofd_date"].apply(lambda x: x[:10])
        # df.dropna(subset=["plan_type"], inplace=True)
        df = df.loc[
            (df["plan_type"] != "weekly-run")
            & (df["station"].isin(nodes))
            & (df["ofd_date"] == ofd_date)
        ]
        latest = df.pivot_table(
            index=["station", "cycle"], values="snapshot_id", aggfunc=max
        ).to_dict(orient="index")

        df["keep"] = df[["station", "cycle", "snapshot_id"]].apply(
            select_latest_forecast, axis=1, snapshot_dict=latest
        )
        df = df.loc[df["keep"] == True]
        df.drop(columns="keep", inplace=True)
        return df
    except Exception:
        return pd.DataFrame(
            columns=[
                "ofd_date",
                "station",
                "plan_type",
                "cycle",
                "service_type",
                "scheduling_service_type",
                "spr",
                "routes_output",
                "shift_length",
                "country",
            ]
        )


def download_flex_data(nodes: list, country: str, ofd_date: str = None):
    def pick_flex_share(vals):
        lmcp, share = vals
        if lmcp != nan and lmcp and not share:
            return lmcp
        return share

    if ofd_date is None:
        ofd_date = datetime.datetime.today().strftime("%Y-%m-%d")
    try:
        df = pd.DataFrame(
            columns=[
                "OFD Date",
                "DS",
                "Cycle",
                "Service Type",
                "Volume Share",
                "SPR",
                "Van Share",
                "Van SPR",
            ]
        )
        try:
            df = pd.read_csv(cache_dir / f"{ofd_date}-flex.csv")
        except (OSError):
            if country.upper() in ["UK", "GB", "IR"]:
                df = pd.read_excel(
                    r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\Scheduling\UK\FlexData\AM_Flex_SA_Inpts_File_UK_extra_durations.xlsm",
                    sheet_name="LogTbl",
                    header=1,
                )

            elif country.upper() in ["DE", "AT"]:
                df = pd.read_excel(
                    r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\Scheduling\DE\FlexData\AM_Flex_SA_Inpts_File_DE.xlsm",
                    sheet_name="LogTbl",
                    header=1,
                )

            elif country.upper() == "ES":
                df = pd.read_excel(
                    r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\Scheduling\ES\FlexData\AM_Flex_SA_Inpts_File_ES.xlsm",
                    sheet_name="LogTbl",
                    header=1,
                )

            else:
                pass

        df["Cycle"] = df["Cycle"].apply(lambda x: str(x).upper())
        df = df.loc[
            (df["DS"].isin(nodes))
            & (df["Service Type"].isin(["Next Day", "Next"]))
            & (df["OFD Date"] == ofd_date)
        ]
        df.fillna("", inplace=True)
        df["Volume Share"] = df[["Share in LMCP", "Volume Share"]].apply(
            pick_flex_share, axis=1
        )
        return df

    except Exception:
        return pd.DataFrame(
            columns=[
                "OFD Date",
                "DS",
                "Cycle",
                "Service Type",
                "Volume Share",
                "SPR",
                "Van Share",
                "Van SPR",
            ]
        )


def download_siphon_volume(nodes: list, ofd_date: str = None):
    if ofd_date is None:
        ofd_date = datetime.datetime.today().strftime("%Y-%m-%d")
    try:
        try:
            df = pd.read_csv(cache_dir / f"{ofd_date}-volume.csv")
        except (OSError):
            df = pd.read_csv(
                r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\AM Shift\6. Handover Report\New Scheduling Model for all countries - Morning report\SIPHON Logs - updated\scheduling_model_inputs.txt",
                sep="\t",
                converters={"ofd_date": str},
            )
        df["ofd_date"] = df["ofd_date"].apply(lambda x: x[:10])
        df.dropna(subset=["plan_type"], inplace=True)
        df = df.loc[
            (df["ofd_date"] == ofd_date)
            & (df["station"].isin(nodes))
            & (df["plan_type"] != "weekly-run")
        ]
        return df
    except Exception:
        return pd.DataFrame(
            columns=[
                "ofd_date",
                "station",
                "cycle",
                "volume_forecast",
                "hz_share",
                "plan_type",
                "country",
            ]
        )


def download_siphon_inputs(nodes: list, ofd_date: str = None):
    if ofd_date is None:
        ofd_date = datetime.datetime.today().strftime("%Y-%m-%d")
    try:
        try:
            df = pd.read_csv(cache_dir / f"{ofd_date}-volume.csv")
        except (OSError):
            df = pd.read_csv(
                r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\AM Shift\6. Handover Report\New Scheduling Model for all countries - Morning report\SIPHON Logs - updated\scheduling_model_inputs.txt",
                sep="\t",
                converters={"ofd_date": str},
            )
        df["ofd_date"] = df["ofd_date"].apply(lambda x: x[:10])
        df = df.loc[
            (df["ofd_date"] == ofd_date)
            & (df["station"].isin(nodes))
            & (df["plan_type"] != "weekly-run")
        ]
        return df
    except Exception:
        return pd.DataFrame(
            columns=[
                "ofd_date",
                "country",
                "station",
                "cycle",
                "volume_forecast",
                "cancellation_policy",
                "flex_share",
                "hz_share",
                "snapshot_id",
                "plan_type",
            ]
        )


def get_flex_data(
    flex_data,
    ds: str,
    cycle: str,
    files_path,
    rlas: bool = False,
    volumes: pd.DataFrame = None,
    volume_override: float = 0,
):
    if volumes is None:
        volumes = pd.DataFrame()
    try:
        if isinstance(flex_data, str):
            df = pd.read_csv(flex_data)
        elif isinstance(flex_data, pd.DataFrame):
            df = flex_data
        else:
            logger.error(
                "Flex data needs to be either pathlike str pointing to CSV or DataFrame"
            )
            raise ValueError
        df = df.loc[(df["DS"] == ds) & (df["Cycle"] == cycle)].copy()
        flex_spr = df["SPR"].values[0]
        if not rlas:
            flex_volume = get_volume(
                ds=ds, cycle=cycle, data=f"{files_path}\\volume_data.csv", output="flex"
            )
        if rlas:
            flex_volume = get_volume(ds=ds, cycle=cycle, data=volumes, output="flex")
        if volume_override:
            flex_volume = volume_override
        flex_routes = flex_volume / flex_spr
        van_values = df.iloc[0].to_dict()
        van_routes = van_values.get("Van Share", float("nan"))
        van_spr = van_values.get("Van SPR", float("nan"))
        return {
            "spr": flex_spr,
            "routes": round(flex_routes, 0),
            "volume": flex_volume,
            "van_spr": van_spr,
            "van_routes": van_routes,
        }
    except Exception:
        return {"spr": nan, "routes": nan, "volume": nan}


def get_volume(ds, cycle, data, output: str = "total"):
    try:
        if isinstance(data, str):
            df = pd.read_csv(data)
        elif isinstance(data, pd.DataFrame):
            df = data
        else:
            print(
                "Volume data needs to be either a string path to CSV file or DataFrame"
            )
            raise ValueError
        df = df.loc[(df["station"] == ds) & (df["cycle"] == cycle)]
        df["hz_volume"] = df["hz_share"] * df["volume_forecast"]
        df["flex_volume"] = df["flex_share"] * df["volume_forecast"]
        plans = df["plan_type"].unique().tolist()
        plans.sort(reverse=True)
        plan = plans[0]
        if output == "total":
            return int(df.loc[df["plan_type"] == plan]["volume_forecast"].values[0])
        elif output == "hz":
            return int(df.loc[df["plan_type"] == plan]["hz_volume"].values[0])
        elif output == "flex":
            return int(df.loc[df["plan_type"] == plan]["flex_volume"].values[0])
        else:
            print('Output should be one of "hz" or "total" or "flex"')
            raise ValueError
    except Exception:
        return nan


def prettify_html(input_html: str) -> str:
    """Formats the tables into easier to read format

    Args:
        input_html (str): html string with tables

    Returns:
        str: Coloured html
    """
    soup = BeautifulSoup(input_html, "lxml")
    colors = {
        "actual_header": "#70ad47",
        "actual_row": "#c6e0b4",
        "actual_row_alt": "#e2efda",
        "forecast_header": "#ed7c31",
        "forecast_row": "#f8cbad",
        "forecast_row_alt": "#fce4d6",
        "index_header": "#4473c4",
        "index_row": "#b4c6e7",
        "index_row_alt": "#d9e1f2",
    }

    tables = soup.find_all("table", class_="cycle")

    for tbl in tables:
        row_num = 1
        for row in tbl.find_all("tr"):
            col_num = 0
            if row.find_all("th"):
                for cell in row.find_all("th"):
                    if col_num < 5:
                        cell.attrs = {"bgcolor": colors["actual_header"]}
                        cell.string.wrap(soup.new_tag("font", color="black"))
                        cell.font.string.wrap(soup.new_tag("b"))
                    elif col_num >= 5:
                        cell.attrs = {"bgcolor": colors["forecast_header"]}
                        cell.string.wrap(soup.new_tag("font", color="black"))
                        cell.font.string.wrap(soup.new_tag("b"))
                    else:
                        cell.attrs = {"bgcolor": colors["index_header"]}
                        cell.string.wrap(soup.new_tag("font", color="black"))
                        cell.font.string.wrap(soup.new_tag("b"))

                    col_num += 1

            elif row_num == len(tbl.find_all("tr")):
                for cell in row.find_all("td"):
                    if col_num < 5:
                        cell.attrs = {"bgcolor": colors["actual_header"]}
                        cell.string.wrap(soup.new_tag("font", color="black"))
                        cell.font.string.wrap(soup.new_tag("b"))
                    elif col_num >= 5:
                        cell.attrs = {"bgcolor": colors["forecast_header"]}
                        cell.string.wrap(soup.new_tag("font", color="black"))
                        cell.font.string.wrap(soup.new_tag("b"))
                    else:
                        cell.attrs = {"bgcolor": colors["index_header"]}
                        cell.string.wrap(soup.new_tag("font", color="black"))
                        cell.font.string.wrap(soup.new_tag("b"))

                    col_num += 1

            else:
                for cell in row.find_all("td"):
                    if col_num < 5:
                        if not row_num % 2:
                            cell.attrs = {"bgcolor": colors["actual_row"]}
                        else:
                            cell.attrs = {"bgcolor": colors["actual_row_alt"]}
                    elif col_num >= 5:
                        if not row_num % 2:
                            cell.attrs = {"bgcolor": colors["forecast_row"]}
                        else:
                            cell.attrs = {"bgcolor": colors["forecast_row_alt"]}
                    else:
                        if not row_num % 2:
                            cell.attrs = {"bgcolor": colors["index_row"]}
                        else:
                            cell.attrs = {"bgcolor": colors["index_row_alt"]}
                    col_num += 1

            row_num += 1

    return str(soup)


def get_route_length(node, cycle, ofd: str = None):
    if not ofd:
        ofd = datetime.date.today().strftime("%Y-%m-%d")
    rcs = amzl_requests.RcsData(node, cycle, ofd, amzl_requests.Auth()).get_inputs(
        float("nan")
    )
    rcs.fillna(float("nan"), inplace=True)
    minutes = []
    for service in rcs.sort_values("Service type", ascending=True)[
        "Service type"
    ].unique():
        scheduled_minutes = rcs.loc[
            (rcs["Service type"] == service) & (~rcs["Scheduled"].isna())
        ]["Shift length"].unique()

        scheduled_minutes = [
            str(f"{mins/60:.1f}h") for mins in scheduled_minutes if not pd.isna(mins)
        ]
        minutes.append(f"{service}: {'/'.join(scheduled_minutes)}")
    return "\n".join(minutes)