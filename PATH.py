import os

desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
app_path = os.path.dirname(os.path.abspath(__file__))

SOURCE_PPT_PATH = os.path.join(app_path, r"오전예배_source.pptx")
OUTPUT_SAVE_PATH = os.path.join(desktop_path, r"오전예배.pptx")

RESPONSIVE_READING_PATH = r"C:\Users\giveroot\Documents\주의길PPT\교독문\주의길_교독문"
HYMN_PATH = r"C:\Users\giveroot\Documents\주의길PPT\새찬송가16x9"

# # test code
# print(SOURCE_PPT_PATH)
# print(OUTPUT_SAVE_PATH)
# print(RESPONSIVE_READING_PATH)
# print(HYMN_PATH)