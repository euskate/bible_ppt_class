import sys
import win32com.client


def open_slide(path):
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    prs = app.Presentations.Open(path)
    return app, prs


def copy_slide_all(prs):
    count = len(prs.Slides)
    prs.Slides.Range(range(1, count + 1)).copy()
    return count


def paste_slide(dst_prs, paste_slide_number):
    dst_prs.Slides.paste(paste_slide_number)


def copy_desgin_slide(src_prs, dst_prs, count, paste_slide_number):
    for i in range(count):
        dst_prs.Slides[paste_slide_number + i].Design = src_prs.Slides[i + 1].Design


# path = "c:\\Users\\Administrator\\Desktop\\WorkSpace\\pyPptx\\new-file-name.pptx"
path = (
    "C:\\Users\\Administrator\\Desktop\\WorkSpace\\pyPptx\\새찬송가16x9\\NHymn016h_Wide.PPT"
)
path2 = (
    "c:\\Users\\Administrator\\Desktop\\WorkSpace\\pyPptx\\오전예배 (16x10)_20201129.pptx"
)
paste_slide_number = 20
desgin_slide_number = paste_slide_number - 1

ppt, prs = open_slide(path)
ppt2, prs2 = open_slide(path2)
count = copy_slide_all(prs)
paste_slide(prs2, paste_slide_number)
copy_desgin_slide(prs, prs2, count, desgin_slide_number)
