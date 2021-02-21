import pickle
import re
from bible_function import (
    abbreviation,
    parse_paragraph,
    dict_contents,
    parsing_contents,
)
from pptx import Presentation
from pptx.util import Pt
from win32com.client import Dispatch
from os.path import abspath, join, dirname

BASE_DIR = dirname(abspath(__file__))


TEMPLATE_PATH = join(BASE_DIR, "verse_template.pptx")
SAVEFILE_NAME = join(BASE_DIR, "new-file-name.pptx")
# BIBLE_DF = join(BASE_DIR, "bible_dataframe.pickle")
# BIBLE_DICT = join(BASE_DIR, "bible_dict.pickle")

## Load pickle
with open("bible_list.pickle", "rb") as fr:
    bible_list = pickle.load(fr)
with open("bible_dict.pickle", "rb") as fr:
    bookDict = pickle.load(fr)

raw = """
【 나와 함께 낙원에 있으리라(눅23:38-43절) 】 

◉2020년 7월 5일 맥추감사주일              ●245장(저 좋은 낙원 이르니)

●마태복음 27장 37-44절.
“37 그 머리 위에 이는 유대인의 왕 예수라 쓴 죄패를 붙였더라 / 이 때에 예수와 함께 강도 둘이 십자가에 못 박히니 하나는 우편에, 하나는 좌편에 있더라 / 39 지나가는 자들은 자기 머리를 흔들며 예수를 모욕하여 / 40 이르되 성전을 헐고 사흘에 짓는 자여. 네가 만일 하나님의 아들이어든 자기를 구원하고 십자가에서 내려오라 하며 / 41 그와 같이 대제사장들도 서기관들과 장로들과 함께 희롱하여 이르되 / 42 그가 남은 구원하였으되 자기는 구원할 수 없도다. 그가 이스라엘의 왕이로다. 지금 십자가에서 내려올지어다. 그리하면 우리가 믿겠노라 / 43 그가 하나님을 신뢰하니 하나님이 원하시면 이제 그를 구원하실지라. 그의 말이 나는 하나님의 아들이라 하였도다 하며 / 44 함께 십자가에 못 박힌 강도들도 이와 같이 욕하더라.”

●마가복음 15장 26-32절.
“26 그 위에 있는 죄패에 유대인의 왕이라 썼고 / 27 강도 둘을 예수와 함께 십자가에 못 박으니 하나는 그의 우편에, 하나는 좌편에 있더라 / 28 (없음) / 29 지나가는 자들은 자기 머리를 흔들며 예수를 모욕하여 이르되, 아하! 성전을 헐고 사흘에 짓는다는 자여 / 30 네가 너를 구원하여 십자가에서 내려오라 하고 / 31 그와 같이 대제사장들도 서기관들과 함게 희롱하며 서로 말하되, 그는 남은 구원하였으되, 자기는 구원할 수 없도다 / 32 이스라엘의 왕 그리스도가 지금 십자가에서 내려와 우리가 보고 믿게 할지어다 하며, 함께 십자가에 못 박힌 자들도 예수를 욕하더라.”

●39-41절.
“39 달린 행악자 중 하나는 비방하여 이르되, 네가 그리스도가 아니냐 너와 우리를 구원하라 하되 / 40 하나는 그 사람을 꾸짖어 이르되, 네가 동일한 정죄를 받고서도 하나님을 두려워하지 아니하느냐 / 41 우리는 우리가 행한 일에 상당한 보응을 받는 것이니 이에 당연하거니와, 이 사람이 행한 것은 옳지 않은 것이 없느니라 하고”

●이사야서 6장 5절. 
“그 때에 내가 말하되 화로다 나여 망하게 되었도다. 나는 입술이 부정한 사람이요, 나는 입술이 부정한 백성 중에 거주하면서 만군의 여호와이신 왕을 뵈었음이로다 하였더라”

●누가복음 5장 8절.
“시몬 베드로가 이를 보고, 예수의 무릎 아래에 엎드려 이르되, 주여 나를 떠나소서. 나는 죄인이로소이다 하니”라고 고백했습니다. 

●41절.
“우리는 우리가 행한 일에 상당한 보응을 받는 것이니 이에 당연하거니와 이 사람이 행한 것은 옳지 않은 것이 없느니라 하고”. 

●40절.
“하나는 그 사람을 꾸짖어 이르되, 네가 동일한 정죄를 받고서도 하나님을 두려워하지 아니하느냐”. 

●42절.
“이르되 예수여! 당신의 나라에 임하실 때에 나를 기억하소서 하니” 

●42절(흠정역 번역). 
“주여, 주께서 주의 왕국으로 들어오실 때에 나를 기억하옵소서”

●마가복음 1장 15절.
“이르시되 때가 찼고, 하나님의 나라가 가까이 왔으니, 회개하고 복음을 믿으라 하시더라.”

●43절.
“내가 진실로 네게 이르노니 오늘 네가 나와 함께 낙원에 있으리라 하시니라”

●마태복음 25장 46절.
“그들은 영벌에, 의인들은 영생에 들어가리라 하시니라”
"""
# PPTX 생성 함수 : 성경 본문(제목포함) + 한글 텍스트 파일
def present(
    resultList,
    template_path,
    save_file_name,
    main_book,
    main_chapter,
    main_verse_start,
    main_verse_end,
    keys,
    contentsDict,
    main_title,
):
    prs = Presentation(template_path)

    # 타이틀 삽입
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "성경봉독"
    subtitle.text = (
        f"{bookDict[main_book]} {main_chapter}장 {main_verse_start}-{main_verse_end}절"
    )
    # slide.shapes[2].text_frame.text = "(신약 p.)"

    # 본문 삽입
    for key in keys:
        contents_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(contents_slide_layout)
        title = slide.shapes.title
        contents = slide.placeholders[1]

        title.text = key
        contents.text = contentsDict[key]

    #############
    # 타이틀 삽입
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "말씀선포"
    subtitle.text = main_title

    for book, chapter, verse_start, verse_end, contents in resultList:
        if not verse_end:
            key = f"{bookDict[book]} {chapter}장 {verse_start}절"
            # pptx
            contents_slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(contents_slide_layout)
            title_placeholder = slide.shapes.title
            contents_placeholder = slide.placeholders[1]

            title_placeholder.text = key
            contents_placeholder.text = contents[0]
        else:
            key = f"{bookDict[book]} {chapter}장 {verse_start}-{verse_end}절"

            contents_slide_layout = prs.slide_layouts[2]
            slide = prs.slides.add_slide(contents_slide_layout)

            title_placeholder = slide.shapes.title
            title_placeholder.text = key

            contents_text_frame = slide.placeholders[1].text_frame
            for _ in range(len(contents) - 1):
                contents_text_frame.add_paragraph()
            for i in range(len(contents)):
                tmp1 = contents_text_frame.paragraphs[i].add_run()
                try:  # 본문에 실수로 구절번호를 생략한 경우 예외처리
                    tmp1.text = re.search("\d+", contents[i]).group()
                except:
                    pass
                tmp1.font.size = Pt(36)
                tmp2 = contents_text_frame.paragraphs[i].add_run()
                tmp2.text = re.sub("\d+", "", contents[i])

    prs.save(save_file_name)


# 파워포인트 열기
def open_ppt(filename=SAVEFILE_NAME):
    PPT = Dispatch("PowerPoint.Application")
    PPT.Visible = True
    PPT.Presentations.Open(abspath(filename))


# 성경 본문만 PPT 만들기
def make_ppt_only_text(paragraph):
    main_book, main_chapter, main_verse_start, main_verse_end = parse_paragraph(
        paragraph
    )
    keys, contentsDict = dict_contents(
        main_book, main_chapter, main_verse_start, main_verse_end
    )
    prs = Presentation(TEMPLATE_PATH)

    for key in keys:
        contents_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(contents_slide_layout)
        title = slide.shapes.title
        contents = slide.placeholders[1]

        title.text = key
        contents.text = contentsDict[key]

    prs.save(SAVEFILE_NAME)

    open_ppt(SAVEFILE_NAME)


def make_ppt_text(raw):
    (
        resultList,
        main_book,
        main_chapter,
        main_verse_start,
        main_verse_end,
        main_title,
    ) = parsing_contents(raw)

    print(main_book, main_chapter, main_verse_start, main_verse_end, main_title)

    keys, contentsDict = dict_contents(
        main_book, main_chapter, main_verse_start, main_verse_end
    )

    # print(resultList)

    present(
        resultList=resultList,
        template_path=TEMPLATE_PATH,
        save_file_name=SAVEFILE_NAME,
        main_book=main_book,
        main_chapter=main_chapter,
        main_verse_start=main_verse_start,
        main_verse_end=main_verse_end,
        keys=keys,
        contentsDict=contentsDict,
        main_title=main_title,
    )

    # open_ppt(SAVEFILE_NAME)


if __name__ == "__main__":
    make_ppt_text(raw)