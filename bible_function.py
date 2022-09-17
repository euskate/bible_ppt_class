import pickle
from re import compile

# # 업데이트 사항
# - 따옴표 그룹 추가
# 편 추가

## Load pickle
with open("bible_list.pickle", "rb") as fr:
    bible_list = pickle.load(fr)
with open("bible_dict.pickle", "rb") as fr:
    bookDict = pickle.load(fr)

# 성경 약자 반환 함수
def abbreviation(book):
    if book in list(bookDict.values()):
        book = list(bookDict.keys())[list(bookDict.values()).index(book)]
    # 성경 권 이름 뒤에 '서'를 붙이는 예외처리
    elif book[-1] == "서" and book[:-1] in list(bookDict.values()):
        book = list(bookDict.keys())[list(bookDict.values()).index(book[:-1])]
    # 성경 권에 '역대기'가 붙는 예외처리 ex) 역대기상 -> 역대상
    elif book[:3] == "역대기":
        book = list(bookDict.keys())[
            list(bookDict.values()).index(book.replace("역대기", "역대"))
        ]
    return book


# 정규식 파싱 함수
def parse_paragraph(paragraph):
    # 장 + 편 추가
    regex = compile(
        "(?P<book>[ㄱ-ㅣ가-힣]+)\s*(?P<chapter>[0-9]+):?장?편?\s*(?P<verse_start>[0-9]+)-*(?P<verse_end>[0-9]+)*절?"
    )
    parsed_paragraph = regex.search(paragraph)
    book = parsed_paragraph.group("book")
    chapter = parsed_paragraph.group("chapter")
    verse_start = parsed_paragraph.group("verse_start")
    verse_end = parsed_paragraph.group("verse_end")

    book = abbreviation(book)

    return book, chapter, verse_start, verse_end


# 검색 구절을 사전형태로 반환하는 함수
def dict_contents(book, chapter, verse_start, verse_end):
    if verse_end:  # 여러 구절일 때
        search_result = [
            b[3]
            for b in bible_list
            if b[0] == book
            and b[1] == int(chapter)
            and b[2] >= int(verse_start)
            and b[2] <= int(verse_end)
        ]
        verses = range(int(verse_start), int(verse_end) + 1)
        if book == "시":  # 시편일 경우 "편"으로 추가
            keys = [f"{bookDict[book]} {chapter}편 {i}절" for i in verses]
        else:
            keys = [f"{bookDict[book]} {chapter}장 {i}절" for i in verses]
    else:  # 한구절만 있을 때
        search_result = [
            b[3]
            for b in bible_list
            if b[0] == book and b[1] == int(chapter) and b[2] == int(verse_start)
        ]
        if book == "시":  # 시편일 경우 "편"으로 추가
            keys = [f"{bookDict[book]} {chapter}편 {verse_start}절"]
        else:
            keys = [f"{bookDict[book]} {chapter}장 {verse_start}절"]
    contentsDict = dict(zip(keys, search_result))
    return keys, contentsDict


# 본문 파싱해주는 함수 Strict(엄격) => (book, chapter, verse_start, verse_end, contents)
def parsing_contents(raw):
    regex = compile(r"【\s*?(?P<title>.*?)\((?P<main_verse>.*?)\)?s*?】")
    main_title = regex.search(raw).group("title").strip()
    main_verse = regex.search(raw).group("main_verse")
    main_book, main_chapter, main_verse_start, main_verse_end = parse_paragraph(
        main_verse
    )

    # 따옴표 "󰡒"(기호 안보임) 추가"
    regex = compile(r"●(.*?)\n ?[“󰡒󰡒](.*?)[”󰡓󰡓]")
    all_p = regex.findall(raw)

    resultList = list()
    for i, c in all_p:
        (book, chapter, verse_start, verse_end) = (None, None, None, None)
        try:
            book, chapter, verse_start, verse_end = parse_paragraph(i)
        except:  # 권, 장이 없고, 절만 있는 경우
            regex = compile("(?P<verse_start>[0-9]+)절?-*(?P<verse_end>[0-9]+)*절?")
            parsed_p = regex.search(i)
            if parsed_p:
                book = main_book
                chapter = main_chapter
                verse_start = parsed_p.group("verse_start")
                verse_end = parsed_p.group("verse_end")
            else:
                pass  # 정규화된 성경구절이 아닌경우
        if verse_end:
            contents = [p.strip() for p in c.split("/")]
        else:
            contents = [c.strip()]
        resultList.append((book, chapter, verse_start, verse_end, contents))
    return (
        resultList,
        main_book,
        main_chapter,
        main_verse_start,
        main_verse_end,
        main_title,
    )


# parsing_contents 중 메인 구절 파싱 부분만
def extract_main_verse(raw):
    regex = compile(r"【\s*?(?P<title>.*?)\((?P<main_verse>.*?)\)?s*?】")
    main_title = regex.search(raw).group("title").strip()
    main_verse = regex.search(raw).group("main_verse")
    main_book, main_chapter, main_verse_start, main_verse_end = parse_paragraph(
        main_verse
    )
    return (main_book, main_chapter, main_verse_start, main_verse_end, main_title)


# parsing_contents 중 비엄격 결과리스트로 반환 (subtitle, text)
def copy_contents(raw):
    # 따옴표 "󰡒"(기호 안보임) 추가"
    regex = compile(r"●(.*?)\..*?\n.*?[“󰡒](.*?)[”󰡓]")
    all_p = regex.findall(raw)

    resultList = list()
    for i, c in all_p:
        (subtitle, text) = (None, None)
        subtitle = i
        if "/" in c:
            text = [p.strip() for p in c.split("/")]
        else:
            text = [c.strip()]
        resultList.append((subtitle, text))
    return resultList