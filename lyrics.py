from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

ppt = Presentation()

# slide dimensions
slide_width = ppt.slide_width.inches
slide_height = ppt.slide_height.inches

# text box dimensions
text_width = Inches(8)
text_height = Inches(2)

# get input
print("가사 입력, 입력 끝날시 q")
text_input = []
while True:
    l = input("")
    if l=='q':
        break
    if l == "":
        continue
    text_input.append(l)

# print("input done ")
# for line in text_input:
#     print(">", line)
# print("done")

line_num = len(text_input)
for idx in range(0, len(text_input), 2):
    text = text_input[idx]
    if idx+1 < line_num:
        text += "\n" + text_input[idx+1]

    slide_layout = ppt.slide_layouts[6]
    slide = ppt.slides.add_slide(slide_layout)

    # slide background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0,0,0)

    # centering text box
    left = (slide_width - text_width.inches) / 2
    top = (slide_height - text_height.inches) / 2

    textBox = slide.shapes.add_textbox(Inches(left), Inches(top), text_width, text_height)
    text_frame = textBox.text_frame
    # text_frame.text = text

    # Style the text
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(36)
    run.font.color.rgb = RGBColor(255, 255, 255)

ppt.save('test.pptx')

print("done")