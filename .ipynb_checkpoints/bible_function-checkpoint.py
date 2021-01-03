import pickle
import re

## Load pickle
with open("bible_dataframe.pickle", "rb") as fr:
    bibleDF = pickle.load(fr)
with open("bible_dict.pickle", "rb") as fr:
    bookDict = pickle.load(fr)

# 성경 약자 반환 함수
def abbreviation(book):
    if book in list(bookDict.values()):
        book = list(bookDict.keys())[list(bookDict.values()).index(book)]
    return book


# 정규식 파싱 함수
def parse_paragraph(paragraph):
    regex = re.compile(
        "(?P<book>[ㄱ-ㅣ가-힣]+)\s*(?P<chapter>[0-9]+):?장?\s*(?P<verse_start>[0-9]+)-*(?P<verse_end>[0-9]+)*절?"
    )
    parsed_paragraph = regex.search(paragraph)
    book = parsed_paragraph.group("book")
    chapter = parsed_paragraph.group("chapter")
    verse_start = parsed_paragraph.group("verse_start")
    verse_end = parsed_paragraph.group("verse_end")

    if book in list(bookDict.values()):
        book = list(bookDict.keys())[list(bookDict.values()).index(book)]

    return book, chapter, verse_start, verse_end


# 검색 구절을 사전형태로 반환하는 함수
def dict_contents(book, chapter, verse_start, verse_end):
    search_result = bibleDF.loc[
        (bibleDF.book == book)
        & (bibleDF.chapter == int(chapter))
        & (bibleDF.verse >= int(verse_start))
        & (bibleDF.verse <= int(verse_end)),
        "contents",
    ]
    verses = range(int(verse_start), int(verse_end) + 1)
    keys = [f"{bookDict[book]} {chapter}장 {i}절" for i in verses]
    contentsDict = dict(zip(keys, search_result))
    return keys, contentsDict


# 본문 파싱해주는 함수
def parsing_contents(raw):
    regex = re.compile(r"【\s*?(?P<title>.*?)\((?P<main_verse>.*?)\)?s*?】")
    main_title = regex.search(raw).group("title").strip()
    main_verse = regex.search(raw).group("main_verse")
    main_book, main_chapter, main_verse_start, main_verse_end = parse_paragraph(
        main_verse
    )

    regex = re.compile(r"●(.*?)\n“(.*?)”")
    all_p = regex.findall(raw)

    resultList = list()
    for i, c in all_p:
        regex = re.compile(
            "(?P<book>[ㄱ-ㅣ가-힣]+)\s*(?P<chapter>[0-9]+):?장?\s*(?P<verse_start>[0-9]+)-*(?P<verse_end>[0-9]+)*절?"
        )
        parsed_paragraph = regex.search(i)
        if parsed_paragraph:
            book = parsed_paragraph.group("book")
            chapter = parsed_paragraph.group("chapter")
            verse_start = parsed_paragraph.group("verse_start")
            verse_end = parsed_paragraph.group("verse_end")
            book = abbreviation(book)
        else:
            regex = re.compile("(?P<verse_start>[0-9]+)절?-*(?P<verse_end>[0-9]+)*절?")
            parsed_p = regex.search(i)
            book = main_book
            chapter = main_chapter
            verse_start = parsed_p.group("verse_start")
            verse_end = parsed_p.group("verse_end")
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
