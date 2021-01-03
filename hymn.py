import sys
import win32com.client


def get_slide_number(path):
    app = win32com.client.Dispatch("PowerPoint.Application")
    prs = app.Presentations.Open(path)
    sec = prs.SectionProperties
    hymn_first_slide = sec.FirstSlide(4)
    hymn_slide_count = sec.SlidesCount(4)
    app.Quit()
    return hymn_first_slide, hymn_slide_count


path = (
    "c:\\Users\\Administrator\\Desktop\\WorkSpace\\pyPptx\\오전예배 (16x10)_20201129.pptx"
)

a, b = get_slide_number(path)

print(a, b)