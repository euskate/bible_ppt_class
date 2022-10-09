# 이전 목표 완성 : 찬양PPT 복사 배경이미지 소스 원본 붙여넣기
# 다음 목표
# 성경봉독 및 말씀선포 복사 붙여넣기 오류 잡기 (찬양 제목 부분)
# 찬양 슬라이드 삽입 위치 정상화

import json
import win32com.client
from time import sleep
import os
from bible_function import (
    abbreviation,
    parse_paragraph,
    dict_contents,
    parsing_contents,
    extract_main_verse,
    copy_contents,
    bookDict,
)
from PATH import SOURCE_PPT_PATH, RESPONSIVE_READING_PATH, HYMN_PATH

SLEEP_TIME = 0.3

class Powerpoint:
    """
    Python PowerPoint win32 object
    """

    # 파워포인트 생성
    def init_app(self):
        self.app = win32com.client.Dispatch("PowerPoint.Application")
        self.app.Visible = True
        self.count = 0  # 전체 슬라이드 수 초기화
        return self.app

    # 프레젠테이션 열기
    def open_prs(self, path):
        self.path = path
        self.prs = self.app.Presentations.Open(self.path)
        return self.prs

    # 프레젠테이션 저장하기
    def save_prs(self, path):
        self.path = path
        self.prs.SaveAs(path)

    # 프레젠테이션 닫기
    def close_prs(self):
        self.prs.Close()

    # 전체 슬라이드 수 세기
    def get_count_slide(self):
        self.count = len(self.prs.Slides)  # 전체 슬라이드 길이를 반환한다

    # 전체 슬라이드 복사
    def copy_slide_all(self):
        self.get_count_slide()
        self.prs.Slides.Range(range(1, self.count + 1)).copy()

    # 슬라이드 복사
    def copy_slide(self, start_slide_number, end_slide_number):
        self.prs.Slides.Range(range(start_slide_number, end_slide_number + 1)).copy()

    # 슬라이드 붙여넣기
    def paste_slide(self, paste_slide_number):
        self.prs.Slides.paste(paste_slide_number)

    # 구역 가져오기
    def get_section_number(self, section_index):
        """[구역 가져오기]
        Args : section_index(int) : 몇번째 구역인지 설정
        Returns:
            [tuple]: (first_slide, section_count)
        """
        sec = self.prs.SectionProperties
        self.first_slide = sec.FirstSlide(section_index)  # 구역 첫번째 페이지
        self.section_count = sec.SlidesCount(section_index)  # 구역 슬라이드 개수
        return self.first_slide, self.section_count

    # 슬라이드 디자인 복붙
    def copy_desgin_slide(self, dst_prs, first_slide_number):
        """[슬라이드 디자인 복붙]

        Args:
            dst_prs ([Presentation_Object]): 붙여넣을 프레젠테이션
            first_slide_number ([int]): 붙여넣기 시작할 슬라이드 번호, self.get_section_number()에서 가져온다.
        """
        self.get_count_slide()  # 복사할 페이지 수를 가져온다.
        for i in range(self.count):
            dst_prs.Slides(first_slide_number + i).Design = self.prs.Slides(
                i + 1
            ).Design

    # 구역 지우기
    def del_section(self, first_slide, section_count):
        self.prs.Slides.Range(range(first_slide, first_slide + section_count)).Delete()

    # 찬양 번호 바꾸기
    def change_hymn_number(self, hymn_number, section_index):
        first_slide, _ = self.get_section_number(section_index)
        self.prs.Slides(first_slide).Shapes(2).Textframe.TextRange.Runs(
            2
        ).Text = hymn_number

    # 전환 애니메이션 적용하기  : main script5
    def change_transition(self, start=1, end=1):
        """
        docstring
        """
        # 화면전환 : 밝기변화
        self.get_count_slide()
        end = self.count

        self.prs.Slides.Range(
            range(start, end)
        ).SlideShowTransition.EntryEffect = 3849  # 밝기변화
        self.prs.Slides.Range(
            range(start, end)
        ).SlideShowTransition.Duration = 0.5  # 시간
        pass

    def get_section(self, section_name):
        section_index = [
            i
            for i in range(1, self.prs.SectionProperties.Count + 1)
            if self.prs.SectionProperties.Name(i) == section_name
        ][0]
        sec = self.prs.SectionProperties
        first_slide = sec.FirstSlide(section_index)  # 구역 첫번째 페이지
        section_count = sec.SlidesCount(section_index)  # 구역 슬라이드 개수
        return first_slide, section_count

    # 제목 레이아웃 부제목 변경
    def change_subtitle(self, slide, subtitle):
        self.prs.Slides(slide).Shapes(2).Textframe.TextRange.Text = subtitle

    # '성경봉독' 섹션 입력
    def input_verse(self, slide, keys, contentsDict):
        for key in keys:
            slide += 1
            self.prs.Slides.AddSlide(
                slide, self.prs.SlideMaster.CustomLayouts(2)
            )  # 슬라이드 추가
            self.prs.Slides(
                slide
            ).Shapes.Title.TextFrame.TextRange.Text = key  # 제목 텍스트 입력
            self.prs.Slides(slide).Shapes(2).TextFrame.TextRange.Text = contentsDict[
                key
            ]  # 내용 텍스트 입력
        pass

    # '말씀 선포' 섹션 입력
    def input_hwp(self, slide, resultList):
        for book, chapter, verse_start, verse_end, contents in resultList:
            slide += 1  # 섹션 시작 슬라이드 다음부터
            if not verse_end:
                key = f"{bookDict[book]} {chapter}장 {verse_start}절"
                # pptx
                self.prs.Slides.AddSlide(slide, self.prs.SlideMaster.CustomLayouts(2))
                self.prs.Slides(
                    slide
                ).Shapes.Title.TextFrame.TextRange.Text = key  # 제목 텍스트
                self.prs.Slides(slide).Shapes(2).TextFrame.TextRange.Text = contents[
                    0
                ]  # 내용 텍스트

            else:
                key = f"{bookDict[book]} {chapter}장 {verse_start}-{verse_end}절"

                self.prs.Slides.AddSlide(slide, self.prs.SlideMaster.CustomLayouts(3))
                self.prs.Slides(
                    slide
                ).Shapes.Title.TextFrame.TextRange.Text = key  # 제목 텍스트
                self.prs.Slides(slide).Shapes(2).TextFrame.TextRange.Text = "\n".join(
                    contents
                )  # 내용 텍스트
                self.prs.Slides(slide).Shapes(2).TextFrame.TextRange.Runs(
                    1
                ).Font.Size = 28  # 구절 폰트 크기 변경
                start_chr = 0
                for _ in range(len(contents) - 1):
                    start_chr = (
                        self.prs.Slides(slide)
                        .Shapes(2)
                        .TextFrame.TextRange.Find("\n", start_chr)
                        .Start
                    )  # 문단 넘김 시작점 찾기
                    self.prs.Slides(slide).Shapes(2).TextFrame.TextRange.Characters(
                        start_chr + 1, 9999
                    ).Runs(
                        1
                    ).Font.Size = 28  # 구절 폰트 크기 변경







### 폐기소스 
    # def executing():  # 실행
    #     # 소스 ppt 열기
    #     path = SOURCE_PPT_PATH
    #     src_ppt = Powerpoint()
    #     src_ppt.init_app()
    #     src_prs = src_ppt.open_prs(path=path)

    #     re_no = input("교독문 번호를 입력하세요: ")
    #     copy_responsive_reading(re_no, src_ppt=src_ppt, section_index=3)

    #     hymn_number = input("첫번째 찬송가 번호를 입력하세요: ")
    #     copy_hymn(hymn_number, src_ppt=src_ppt, section_index=4)

    #     hymn_number = input("두번째 찬송가 번호를 입력하세요: ")
    #     copy_hymn(hymn_number, src_ppt=src_ppt, section_index=8)

    #     sleep(SLEEP_TIME)

    #     src_ppt.change_transition()


    # if __name__ == "__main__":
    #     executing()