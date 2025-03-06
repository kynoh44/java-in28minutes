import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import os


def parse_udemy_course(html_content, course_title):
    soup = BeautifulSoup(html_content, "html.parser")

    results = []
    section_titles = soup.find_all("span", class_="section--section-title--svpHP")

    current_section = None
    item_title = None  # 초기화 추가 ✅

    for element in soup.find_all():
        if (
            element.name == "span"
            and element.get("class")
            and "section--section-title--svpHP" in element.get("class")
        ):
            current_section = element.text.strip()

        elif (
            element.name == "span"
            and element.get("class")
            and "section--item-title--EWIuI" in element.get("class")
        ):
            if current_section:
                item_title = element.text.strip()

        if (
            element.name == "span"
            and element.get("class")
            and "section--hidden-on-mobile---ITMr"
            and "section--item-content-summary--Aq9em" in element.get("class")
        ):
            if current_section:
                time_element = element.text.strip()
                results.append(
                    {
                        "강의 제목": course_title,
                        "섹션 타이틀": current_section,
                        "섹션 강의 아이템": (
                            item_title if item_title else "제목 없음"
                        ),  # 기본값 설정 ✅
                        "시간": time_element,
                    }
                )

    return pd.DataFrame(results)


def save_to_excel(df, filename="udemy_course_analysis.xlsx"):
    """
    데이터프레임을 Excel 파일로 저장

    Args:
                    df (pandas.DataFrame): 저장할 데이터프레임
                    filename (str): 저장할 파일 이름
    """
    df.to_excel(filename, index=False)
    print(f"데이터가 {filename}에 성공적으로 저장되었습니다.")


def main():
    # 사용 예시
    course_title = input("Udemy 강의 제목을 입력하세요: ")

    # HTML 파일 경로 입력 또는 직접 HTML 내용 붙여넣기 옵션
    option = input(
        "HTML 파일 경로를 입력하려면 1, URL을 입력하려면 2, HTML 내용을 직접 붙여넣으려면 3을 입력하세요: "
    )

    if option == "1":
        file_path = input("HTML 파일 경로를 입력하세요: ")
        with open(file_path, "r", encoding="utf-8") as file:
            html_content = file.read()
    elif option == "2":
        url = input("Udemy 강의 URL을 입력하세요: ")
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(url, headers=headers)
        html_content = response.text
    else:
        print("HTML 내용을 붙여넣고 마지막에 'END_HTML'을 입력하세요:")
        html_lines = []
        while True:
            line = input()
            if line == "END_HTML":
                break
            html_lines.append(line)
        html_content = "\n".join(html_lines)

    # 데이터 파싱
    df = parse_udemy_course(html_content, course_title)

    # 결과 확인
    print(f"총 {len(df)}개의 강의 아이템이 파싱되었습니다.")
    if not df.empty:
        print("파싱된 데이터 샘플:")
        print(df.head())
    else:
        print("데이터 파싱에 실패했습니다. HTML 구조를 확인해주세요.")

    # Excel 파일로 저장
    if not df.empty:
        output_filename = f"{course_title.replace(' ', '-')}.xlsx"
        save_to_excel(df, output_filename)


if __name__ == "__main__":
    main()
