import os
import shutil
from datetime import datetime, timedelta

import requests
from bs4 import BeautifulSoup

FACULTY_URL = "https://vsu.by/studentam/raspisanie-zanyatij.html"
SCHEDULE_URL = "https://vsu.by{faculty_url}"


def _get_start_of_weekday() -> list[str]:
    now = datetime.now()
    weekday = now.weekday()
    if 0 <= weekday <= 3:
        monday = now - timedelta(days=weekday)
    else:
        monday = now + timedelta(days=(7 - weekday))
    return monday.date().strftime("%d.%m").split(".")


def _find_faculties() -> tuple[list[str], list[str]]:
    try:
        url = FACULTY_URL
        r = requests.get(url)
        soup = BeautifulSoup(r.content, "html.parser")
        links, links_text = [], []
        for link in soup.find_all("a"):
            href = link.get("href")
            if (
                    href
                    and "/universitet/fakultety" in href
                    and "/raspisanie.html" in href
                    and "obucheniya-inostrannykh-grazhdan" not in href
            ):
                links_text.append(link.text)
                links.append(href)
        return links, links_text
    except Exception as exc:
        error = f"\n[?]_find_faculties: {exc}"
        print(error)
        input("Press enter to exit\n")
        exit(1)


def _find_schedules(faculty_url: str, monday: list[str]) -> tuple[list[str], list[str]]:
    try:
        url = SCHEDULE_URL.format(faculty_url=faculty_url)
        r = requests.get(url)
        soup = BeautifulSoup(r.content, "html.parser")
        links, links_text = [], []
        for link in soup.find_all("a"):
            href = link.get("href")
            if (
                    href
                    and (".xlsx" in link.text or ".xlsx" in href)
                    and "Расписание" in link.text
                    and (
                    ".".join(monday) in link.text
                    or (
                            monday[1][0] == "0"
                            and f"{monday[0]}.{monday[1][1]}" in link.text
                    )
                    or (
                            monday[0][0] == "0"
                            and monday[1][0] == "0"
                            and f"{monday[0][1]}.{monday[1][1]}" in link.text
                    )
                    or (
                            monday[0][0] == "0"
                            and f"{monday[0][1]}.{monday[1]}" in link.text
                    )
                    # or (
                    #         link.text.index(monday[0]) < link.text.index(monday[1])
                    # )
            )
                    and "ЗФПО" not in href
                    and "ЗФО" not in href
                    and "зф" not in href
                    and "зач" not in link.text
                    and "экзаменов" not in link.text
                    and " к " not in link.text
            ):
                links_text.append(link.text)
                links.append(href)
        return links, links_text
    except Exception as exc:
        error = f"\n[?]_find_schedules: {exc}"
        print(error)
        input("Press enter to exit\n")
        exit(1)


def _download_schedule(link: str, faculty_path: str, schedule_name: str) -> None:
    try:
        url = SCHEDULE_URL.format(faculty_url=link)
        response = requests.get(url, allow_redirects=True)
        schedule_path = os.path.join(faculty_path, f"{schedule_name}.xlsx")
        print(f"[+]Successfully downloaded: {schedule_path}")
        with open(schedule_path, "wb") as f:
            f.write(response.content)
    except Exception as exc:
        error = f"\n[?]_download_schedule: {exc}"
        print(error)


def download():
    try:
        data_folder = os.path.join("data")
        if os.path.exists(data_folder):
            shutil.rmtree(data_folder)
        os.makedirs(data_folder)

        monday = _get_start_of_weekday()
        faculties_links, faculties_name = _find_faculties()

        for faculty_url, faculty_name in zip(faculties_links, faculties_name):
            faculty_name: str = faculty_name.replace("\n", "").replace("\t", "")
            faculty_path = os.path.join(data_folder, faculty_name)
            if not os.path.exists(faculty_path):
                os.makedirs(faculty_path)

            schedules_links, schedules_name = _find_schedules(faculty_url, monday)
            for schedule_link, schedule_name in zip(schedules_links, schedules_name):
                _download_schedule(schedule_link, faculty_path, schedule_name)
    except Exception as exc:
        error = f"\n[?]download: {exc}"
        print(error)
        input("Press enter to exit\n")
        exit(1)


if __name__ == "__main__":
    download()
