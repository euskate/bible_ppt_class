# 찬양 복사 함수 : main script2, script4

from time import sleep
from PATH import HYMN_PATH
from bible_class import Powerpoint

# 대기 시간 설정 (초)
SLEEP_TIME = 0.3

def copy_hymn(hymn_number, src_ppt, section_index):
    """
    찬양 복사 함수
        section_index : 구역순서 { '찬송 1' : 4,  '찬송2' : 8 }
        hymn_number : 새찬송가 번호
        src_ppt : 소스 ppt
    """
    # 찬송가 PPT 파일 경로 확인 및 열기
    hymn_path = f"{HYMN_PATH}\\NHymn{str(hymn_number).zfill(3)}h_Wide.PPT"
    hymn_ppt = Powerpoint()
    hymn_ppt.init_app()
    hymn_ppt.open_prs(hymn_path)
    # 대기 시간 후 전체 슬라이드 복사
    sleep(SLEEP_TIME)
    hymn_ppt.copy_slide_all()
    # 대기 후 소스 PPT 위치 확인 후 이전 내용 지우기
    sleep(SLEEP_TIME)
    first_slide, section_count = src_ppt.get_section_number(section_index)
    try:
      src_ppt.del_section(first_slide + 1, section_count - 2)
    except:
      print("이전가 삭제 실패 예외 발생")
      pass
    # 소스 PPT 활성화하여 원본소스 붙여넣기
    sleep(SLEEP_TIME)
    src_ppt.app.Windows(2).View.GotoSlide(first_slide)
    src_ppt.app.Windows(2).Activate()  # 이전 창(소스ppt) 활성화
    src_ppt.app.CommandBars.ExecuteMso("PasteSourceFormatting")  # 원본소스유지 붙여넣기
    # 찬양 맨 앞장 번호 바꾸기
    sleep(SLEEP_TIME)
    src_ppt.change_hymn_number(hymn_number, section_index)
    # 찬양 PPT 닫기
    hymn_ppt.prs.Close()

    # 폐기 소스
    # win_number = src_ppt.app.Windows.Count  # 현재 ppt 창 번호 따기
    # print("win_number", win_number)
    # src_ppt.paste_slide(first_slide + 1)
    # hymn_ppt.copy_desgin_slide(src_ppt.prs, first_slide + 1)
