# 날짜를 입력하는 method

from ast import main
import datetime

# 날짜가 있는 슬라이드 번호
DATE_SLIDE_NUMBER = 2

def change_date(src_ppt):
    today = datetime.date.today()
    # 날짜에 맞는 텍스트
    new_text = f"{today.month}월 {today.day}일 주일오전예배"
    # 기존 구역 확인 후 삽입
    src_ppt.prs.Slides(DATE_SLIDE_NUMBER).Shapes(1).Textframe.TextRange.Text = new_text
    return


# 테스트 코드
# from bible_class import Powerpoint
# from PATH import SOURCE_PPT_PATH

# if __name__ == "__main__":
#     # 테스트 열기
#     src_ppt = Powerpoint()
#     src_ppt.init_app()
#     src_prs = src_ppt.open_prs(SOURCE_PPT_PATH)
#     # 테스트 메소드 실행
#     change_date(src_ppt)
#     # 테스트 닫기
#     src_ppt.close_prs()