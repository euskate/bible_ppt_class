# 본문 입력 : main script3
# 성경봉독과 한글입력을 구분해야하는 이슈가 있음

from bible_function import (
    dict_contents,
    parsing_contents,
    extract_main_verse,
    bookDict,
)

def copy_text(raw, src_ppt):
    # 본문 권, 장, 절, 제목 추출 extract_main_verse
    (
        main_book,
        main_chapter,
        main_verse_start,
        main_verse_end,
        main_title,
    ) = extract_main_verse(raw)
    # 해당 내용을 딕셔너리 형태로 반환
    keys, contentsDict = dict_contents(
        main_book, main_chapter, main_verse_start, main_verse_end
    )
    # 소스 PPT에서 성경봉독 삽입 위치 찾기 > 이전 내용 지우기 > 부제목변경 > 삽입
    first_slide, section_count = src_ppt.get_section("성경봉독")
    src_ppt.del_section(first_slide + 1, section_count - 2)
    src_ppt.change_subtitle(
        first_slide,
        f"{bookDict[main_book]} {main_chapter}장 {main_verse_start}-{main_verse_end}절",
    )
    src_ppt.input_verse(first_slide, keys, contentsDict)
    (
        resultList,
        main_book,
        main_chapter,
        main_verse_start,
        main_verse_end,
        main_title,
    ) = parsing_contents(raw)
    # 소스 PPT에서 말씀선포 삽입 위치 찾기 > 이전 내용 지우기 > 부제목 변경
    first_slide, section_count = src_ppt.get_section("말씀 선포")
    src_ppt.del_section(first_slide + 1, section_count - 2)
    src_ppt.change_subtitle(first_slide, main_title)
    # 한글내용 입력
    src_ppt.input_hwp(first_slide, resultList)