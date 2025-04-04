import os
import requests
import xlwt
from datetime import datetime, timedelta
from typing import List
from bs4 import BeautifulSoup

group_lessons_list = []

def download_page(url, filename=""):
    folder = os.path.dirname(filename)
    if folder:
        os.makedirs(folder, exist_ok=True)

    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "lxml")
        name = soup.find_all("option", attrs={"selected": "selected"})
        name1 = name[0].text if name else os.path.basename(filename)
        file_path = os.path.join(folder, f"{name1}.html") if folder else f"{name1}.html"
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(response.text)
        return file_path
    else:
        print(f"Ошибка {response.status_code} при загрузке {url}")
        return None

def get_group_page_urls_parse_page(filename="page.html") -> List[str]:
    links_list = []
    with open(filename, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "lxml")
        links = soup.find_all("a")
        for link in links:
            if link.text.strip() == "Посмотреть":
                links_list.append(link.get("href"))
    return links_list

def get_link_page(links_list, day, folder="group_page"):
    os.makedirs(folder, exist_ok=True)
    check_new_week = day.weekday() == 0
    for i, link in enumerate(links_list):
        full_url = f"https://ies.unitech-mo.ru{link}"
        if check_new_week:
            begin_of_week = day.strftime("%d.%m.%Y")
            end_of_week = (day + timedelta(days=6)).strftime("%d.%m.%Y")
            full_url += f"&d={begin_of_week}+-+{end_of_week}"

        filename = os.path.join(folder, f"{i}")
        if download_page(full_url, filename):
            group_name, table = page_parser(f"{filename}.html")
            if table:
                table_parser(group_name, table,i,len(links_list))
                print(f"{i+1}/{len(links_list)}")

def table_parser(group_name, table,i,b):
    tbody = table.find("tbody")
    tr_list = tbody.find_all("tr") if tbody else []
    for tr in tr_list:
        tdlist = tr.find_all("td")
        if len(tdlist) < 6 or tdlist[2].text.strip() == "" or tdlist[1].text == ": - :":
            continue
        num, time, desc, audience, tutor, comments = tdlist[0].text, tdlist[1].text, tdlist[2].text, tdlist[3].text, tdlist[4].text, tdlist[5].text
        if "на самостоятельное обучение" in comments or "дистанционном формате" in comments or "дистанционной форме" in comments:
            audience = "C/Р"
        group_lessons_list.append([group_name, num, time, desc, tutor, audience, comments])
        #print(f"У группы {group_name} найдена {num} пара {tomorrow.strftime('%d.%m.%Y')}, {comments}")

def page_parser(filename):
    with open(filename, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "lxml")
        group_name = soup.find("h2").text.split(" ")[1] if soup.find("h2") else "Неизвестная группа"
        h2_list = soup.find_all("h2", {"class": "text-center"})
        current_h2 = next((h2 for h2 in h2_list if tomorrow.strftime("%d.%m.%Y") in h2.text), None)
        if current_h2:
            div = current_h2.find_next("div", class_="adopt_area_scrollable")
            if div:
                return group_name, div.find("table")
    return group_name, None

def save_to_xls(data, filename="file.xls"):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Расписание")
    headers = ["Группа", "№ пары", "Время", "Дисциплина", "Преподаватель", "Аудитория", "Комментарий"]
    for col, header in enumerate(headers):
        sheet.write(0, col, header)
    for row, entry in enumerate(data, start=1):
        for col, value in enumerate(entry):
            sheet.write(row, col, value)
    workbook.save(filename)
    print(f"Файл {filename} успешно сохранен!")


if __name__ == '__main__':
    day = datetime.now()
    tomorrow = day + timedelta(days=1)
    page_names = []
    links = []
    university_links = [
        "https://ies.unitech-mo.ru/schedule_list_groups?i=19&f=0&k=0",
        "https://ies.unitech-mo.ru/schedule_list_groups?i=20&f=0&k=0",
        "https://ies.unitech-mo.ru/schedule_list_groups?i=21&f=0&k=0",
        "https://ies.unitech-mo.ru/schedule_list_groups?i=22&f=0&k=0"
    ]

    for link in university_links:
        page_name = download_page(link)
        if page_name:
            page_names.append(page_name)
    for name in page_names:
        links.extend(get_group_page_urls_parse_page(name))
    get_link_page(links, tomorrow,folder="university_pages")
    save_to_xls(group_lessons_list, "универ.xls")
    group_lessons_list = []
    url = "https://ies.unitech-mo.ru/schedule_list_groups?i=1&f=0&k=0"
    name = download_page(url)  # Скачиваем страницу
    links = get_group_page_urls_parse_page(name)
    get_link_page(links, tomorrow)
    save_to_xls(group_lessons_list,"колледж.xls")
    print("done")
