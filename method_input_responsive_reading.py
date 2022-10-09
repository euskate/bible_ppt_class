# 교독문 복사 함수 : main script1
import json

LAYOUT_NUMBER_TEXT = 12
LAYOUT_NUMBER_TOGETHER = 13

def input_responsive_reading(re_no, src_ppt, section_index=3):
    """
    교독문 추가 함수
        re_no : 교독문번호
        src_ppt : 소스 PPT
        section_index : 구역순서 { '교독문' : 3  }
    """
    # 데이터 불러오기
    with open("readings.json", "r", encoding="UTF-8") as f:
        reading = json.load(f)[int(re_no)-1]
    heading = reading.get("heading")
    text = reading.get("text")
    # 기존 교독문 구역 확인
    first_slide, section_count = src_ppt.get_section_number(section_index)
    # 교독문 구역 삭제
    src_ppt.del_section(first_slide + 1, section_count - 1)
    # 변수 설정
    prs = src_ppt.prs
    # 교독문 제목 입력
    prs.Slides(first_slide).Shapes(1).Textframe.TextRange.Text = f"교독문 {re_no}번"
    prs.Slides(first_slide).Shapes(2).Textframe.TextRange.Text = heading
    # 추가할 슬라이드별 나눔 및 페이지 세기
    texts = text.split('\n')
    has_together = texts[-1][:5] == "(다같이)"
    page_count = int(len(texts)/2)
    # 교독문 슬라이드 추가
    slide_number = first_slide
    for i in range(page_count):
        slide_number += 1
        prs.Slides.AddSlide(slide_number, prs.SlideMaster.CustomLayouts(LAYOUT_NUMBER_TEXT))
        prs.Slides(slide_number).Shapes(1).TextFrame.TextRange.Text = texts[2*i]
        prs.Slides(slide_number).Shapes(2).TextFrame.TextRange.Text = texts[2*i+1]
    # 다같이 페이지 생성
    if has_together:
        slide_number += 1
        prs.Slides.AddSlide(slide_number, prs.SlideMaster.CustomLayouts(LAYOUT_NUMBER_TOGETHER))
        prs.Slides(slide_number).Shapes(1).TextFrame.TextRange.Text = texts[-1][6:]
        # 아멘 추가
        prs.Slides(slide_number).Shapes(1).TextFrame.TextRange.Text += ".  아멘."
    else:
        prs.Slides(slide_number).Shapes(2).TextFrame.TextRange.Text += ".  아멘."