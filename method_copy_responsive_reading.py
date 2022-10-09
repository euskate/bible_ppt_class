# input_responsive_reading 메소드의 개발로
# 해당 copy 메소드는 사용하지 않습니다.
# (미사용) 교독문 복사 함수 : main script1

from time import sleep
from PATH import RESPONSIVE_READING_PATH
from bible_class import Powerpoint

# 대기 시간 설정 (초)
SLEEP_TIME = 0.3

def copy_responsive_reading(re_no, src_ppt, section_index=3):
    """
    교독문 복사 함수
        section_index : 구역순서 { '교독문' : 3  }
    """


    f_re_no = int(re_no)
    file_name = f"주의길_새교독문{f_re_no:03d}번.pptx"

    path = f"{RESPONSIVE_READING_PATH}\\{file_name}"

    ppt = Powerpoint()
    ppt.init_app()
    ppt.open_prs(path)
    ppt.copy_slide_all()
    first_slide, section_count = src_ppt.get_section_number(section_index)

    # print(first_slide)
    sleep(SLEEP_TIME)
    src_ppt.del_section(first_slide + 1, section_count - 1)
    ### 복사 붙여넣기
    sleep(SLEEP_TIME)
    print("src_ppt 슬라이드 삽입위치", first_slide)
    src_ppt.app.Windows(1).View.GotoSlide(first_slide)
    src_ppt.app.Windows(1).Activate()  # 이전 창(소스ppt) 활성화
    src_ppt.app.CommandBars.ExecuteMso("PasteSourceFormatting")  # 원본소스유지 붙여넣기
    ###
    # src_ppt.paste_slide(first_slide + 1)
    # ppt.copy_desgin_slide(src_ppt.prs, first_slide + 1)
    ppt.prs.Close()
