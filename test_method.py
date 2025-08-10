import unittest
from unittest.mock import MagicMock, patch
from method_copy_text import copy_text
from bible_function import dict_contents

# filepath: c:\b\bible_ppt_class\test_method.py

# copy_text가 method_copy_text.py에 있다고 가정

class TestCopyText(unittest.TestCase):
    @patch("method_copy_text.extract_main_verse")
    @patch("method_copy_text.dict_contents")
    @patch("method_copy_text.parsing_contents")
    @patch("method_copy_text.bookDict", {"시편": "시", "창세기": "창"})
    def test_copy_text_psalm(self, mock_parsing, mock_dict, mock_extract):
        # Arrange
        raw = "시편 23:1-2 제목"
        src_ppt = MagicMock()
        # extract_main_verse 반환값
        mock_extract.return_value = ("시편", 23, 1, 2, "제목")
        # dict_contents 반환값
        mock_dict.return_value = (["23:1", "23:2"], {"23:1": "여호와는 나의 목자시니", "23:2": "그가 나를 누이시며"})
        # parsing_contents 반환값
        mock_parsing.return_value = (
            ["여호와는 나의 목자시니", "그가 나를 누이시며"],
            "시편", 23, 1, 2, "제목"
        )
        src_ppt.get_section.side_effect = [ (1, 3), (5, 2) ]  # 성경봉독, 말씀 선포
        # Act
        copy_text(raw, src_ppt)
        # Assert
        src_ppt.del_section.assert_any_call(2, 1)
        src_ppt.change_subtitle.assert_any_call(1, "시 23편 1-2절")
        src_ppt.input_verse.assert_called_once_with(1, ["23:1", "23:2"], {"23:1": "여호와는 나의 목자시니", "23:2": "그가 나를 누이시며"})
        src_ppt.change_subtitle.assert_any_call(5, "제목")
        src_ppt.input_hwp.assert_called_once_with(5, ["여호와는 나의 목자시니", "그가 나를 누이시며"])

    @patch("method_copy_text.extract_main_verse")
    @patch("method_copy_text.dict_contents")
    @patch("method_copy_text.parsing_contents")
    @patch("method_copy_text.bookDict", {"시편": "시", "창세기": "창"})
    def test_copy_text_other_book(self, mock_parsing, mock_dict, mock_extract):
        raw = "창세기 1:1-3 태초에"
        src_ppt = MagicMock()
        mock_extract.return_value = ("창세기", 1, 1, 3, "태초에")
        mock_dict.return_value = (["1:1", "1:2", "1:3"], {"1:1": "태초에", "1:2": "하나님이", "1:3": "천지를"})
        mock_parsing.return_value = (
            ["태초에", "하나님이", "천지를"],
            "창세기", 1, 1, 3, "태초에"
        )
        src_ppt.get_section.side_effect = [ (2, 4), (7, 3) ]
        copy_text(raw, src_ppt)
        src_ppt.del_section.assert_any_call(3, 2)
        src_ppt.change_subtitle.assert_any_call(2, "창 1장 1-3절")
        src_ppt.input_verse.assert_called_once_with(2, ["1:1", "1:2", "1:3"], {"1:1": "태초에", "1:2": "하나님이", "1:3": "천지를"})
        src_ppt.change_subtitle.assert_any_call(7, "태초에")
        src_ppt.input_hwp.assert_called_once_with(7, ["태초에", "하나님이", "천지를"])


def test_dict_contents():
    keys, contentsDict = dict_contents("시",13,1,6)
    print(keys)
    print(contentsDict)

def test_extract_main_verse():
    from bible_function import extract_main_verse
    raw = " 【영적 침체에서 벗어나기(시13:1-6절)】 "
    print(raw)
    main_book, main_chapter, main_verse_start, main_verse_end, main_title = extract_main_verse(raw)

    print(main_book, main_chapter, main_verse_start, main_verse_end, main_title)

def test_copy_text():
    from method_copy_text import copy_text
    raw = " 【영적 침체에서 벗어나기(시13:1-6절)】 "
    src_ppt = MagicMock()
    src_ppt.get_section.side_effect = [(1, 3), (5, 2)]  # 성경봉독, 말씀 선포
    copy_text(raw, src_ppt)

if __name__ == "__main__":
    # unittest.main()
    # test_dict_contents()
    # test_extract_main_verse()
    test_copy_text()