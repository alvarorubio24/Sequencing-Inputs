import datetime
import os
import time
from pathlib import Path
from threading import Thread

import pandas as pd
from loguru import logger

# globals
cache_dir = Path(f"{os.getcwd()}\\cache\\")
siphon_logs_dir = Path(
    "\\\\ant\\dept-eu\\TBA\\UK\\Business Analyses\\CentralOPS\\AM Shift\\6. Handover Report\\New Scheduling Model for all countries - Morning report\\SIPHON Logs"
)
siphon_logs_updated_dir = Path(
    "\\\\ant\\dept-eu\\TBA\\UK\\Business Analyses\\CentralOPS\\AM Shift\\6. Handover Report\\New Scheduling Model for all countries - Morning report\\SIPHON Logs - updated"
)
queries_dir = Path("\\\\ant\\dept-eu\\TBA\\UK\\Business Analyses\\CentralOPS\\Queries")
today = datetime.datetime.now().strftime("%Y-%m-%d")
ofd_today = today + " 00:00:00"
os.makedirs(cache_dir, exist_ok=True)
logger.add(
    "logs\\cache_data.log",
    rotation="10 MB",
    enqueue=True,
    format="{time:YYYY-MM-DD at HH:mm:ss} | {level} | {message}",
    level=30,
)


for file in [f for f in os.listdir(cache_dir) if not f.startswith(today)]:
    try:
        os.remove(cache_dir / file)
    except OSError:
        pass
    except Exception:
        logger.exception("Uncaught exception happened while removing old cache")


def cache_volume_forecast():
    try:
        logger.info("Caching volume")
        df = pd.read_csv(
            siphon_logs_updated_dir / "scheduling_model_inputs.txt", sep="\t"
        )
        df = df.loc[df["ofd_date"] == ofd_today]
        df.to_csv(cache_dir / (today + "-volume.csv"), index=False)
        logger.info("Caching volume complete!")
    except Exception:
        logger.exception("Uncaught exception encountered while caching volume forecast")


def cache_spr_forecast():
    try:
        logger.info("Caching SPR forecast")
        df = pd.read_csv(
            siphon_logs_updated_dir / "siphon_pull_both_service_types.txt", sep="\t"
        )
        df = df.loc[df["ofd_date"] == ofd_today]
        df.to_csv(cache_dir / (today + "-spr.csv"), index=False)
        logger.info("Caching SPR forecast complete!")
    except Exception:
        logger.exception("Uncaught exception encountered while caching SPR forecast")


def cache_flex_data():
    try:
        logger.info("Caching Flex data")
        # UK flex
        uk = pd.read_excel(
            r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\Scheduling\UK\FlexData\AM_Flex_SA_Inpts_File_UK_extra_durations.xlsm",
            sheet_name="LogTbl",
            header=1,
        )

        # DE flex
        de = pd.read_excel(
            r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\Scheduling\DE\FlexData\AM_Flex_SA_Inpts_File_DE.xlsm",
            sheet_name="LogTbl",
            header=1,
        )

        # ES flex
        es = pd.read_excel(
            r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\Scheduling\ES\FlexData\AM_Flex_SA_Inpts_File_ES.xlsm",
            sheet_name="LogTbl",
            header=1,
        )

        # FR flex
        fr = pd.read_excel(
            r"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\Scheduling\FR\Flexdata\FR_Flex_SA_Inpts_File.xlsm",
            sheet_name="LogTbl",
        )

        out = pd.concat([uk,de,es,fr])
        out = out.loc[out["OFD Date"] == today]
        out.to_csv(cache_dir / (today + "-flex.csv"), index=False)
        logger.info("Caching Flex data complete!")
    except Exception:
        logger.exception("Uncaught exception encountered while caching Flex forecast")


def cache_precuts_data():
    try:
        logger.info("Caching precuts data")
        df = pd.read_csv(siphon_logs_dir / "Siphon_pull_per_DSP.txt", sep="\t")
        df = df.loc[df["ofd_date"] == ofd_today]
        df.to_csv(cache_dir / (today + "-precuts.csv"), index=False)
        logger.info("Caching precuts data complete!")
    except Exception:
        logger.exception("Uncaught exception encountered while caching precuts data")


def cache_past_data():
    try:
        logger.info("Caching historical data")
        df = pd.read_csv(queries_dir / "PlanningMinutesDS_EU.txt", sep="\t")
        df.to_csv(cache_dir / (today + "-historical.csv"), index=False)
        logger.info("Caching historical data complete!")
    except Exception:
        logger.exception("Uncaught exception encountered while caching historical data")


def main():
    threads = []
    funcs = [
        cache_volume_forecast,
        cache_past_data,
        cache_flex_data,
        cache_precuts_data,
        cache_spr_forecast,
    ]

    for func in funcs:
        t = Thread(target=func)
        t.start()
        threads.append(t)

    for thread in threads:
        thread.join()

    os.system(
        f'''copy /Y "\\\\ant\\dept-eu\\TBA\\UK\\Business Analyses\\CentralOPS\\AM Shift\\1. Station Checklists\\1. Checklist Manager\\New Checklist\\service_types.csv" "{cache_dir/'service_types.csv'}"'''
    )

    logger.success("All data cached successfully!")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logger.exception("Uncaught exception encountered")
        input()
    else:
        time.sleep(5)