import tkinter as tk
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor


# # get input
# print("가사 입력, 입력 끝날시 q")
# text_input = []
# while True:
#     l = input("")
#     if l=='q':
#         break
#     if l == "":
#         continue
#     text_input.append(l)



def create_presentation():
    print("create_presentation")
    # get text input
    text_input = text_entry.get("1.0", tk.END).strip()
    lines = text_input.splitlines()
    text_input = []
    for l in lines:
        if not l == "":
            text_input.append(l)

    if not text_input:
        messagebox.showwarning("Input Error", "Please enter text for the slide.")
        return
    
    # get font size
    try:
        font_size = int(font_entry.get())
    except ValueError:
        messagebox.showwarning("Input Error", "Font size has to be an integer.")

    for idx in range(0, len(text_input), 2):
        text = text_input[idx]
        if idx+1 < len(text_input):
            text += "\n" + text_input[idx+1]
        create_slide(text, font_size)

    ppt.save('test.pptx')
    print("done")

def create_slide(text, font_size):
    print('create_slide')
    slide_layout = ppt.slide_layouts[6]
    slide = ppt.slides.add_slide(slide_layout)

    # slide dimensions
    slide_width = ppt.slide_width.inches
    slide_height = ppt.slide_height.inches

    # text box dimensions
    text_width = Inches(8)
    text_height = Inches(2)
    
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
    # center vertically
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    # text_frame.text = text

    # Style the text
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = text # 
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(255, 255, 255)


ppt = Presentation()

gui = tk.Tk()
gui.title("Slide Maker")

tk.Label(gui, text="Enter Lyrics").pack(pady=10)
text_entry = tk.Text(gui, width=50, height=20)
text_entry.pack(pady=5)

# font size
tk.Label(gui, text="Font Size:").pack(pady=10)
font_default = 36
font_entry = tk.Entry(gui)
font_entry.insert(0, str(font_default))
font_entry.pack(pady=5)

tk.Button(gui, text="Create Slide", command=create_presentation).pack(pady=20)
gui.mainloop()

# line_num = len(text_input)
# for idx in range(0, len(text_input), 2):
#     text = text_input[idx]
#     if idx+1 < line_num:
#         text += "\n" + text_input[idx+1]

#     slide_layout = ppt.slide_layouts[6]
#     slide = ppt.slides.add_slide(slide_layout)

#     # slide background
#     background = slide.background
#     fill = background.fill
#     fill.solid()
#     fill.fore_color.rgb = RGBColor(0,0,0)

#     # centering text box
#     left = (slide_width - text_width.inches) / 2
#     top = (slide_height - text_height.inches) / 2

#     textBox = slide.shapes.add_textbox(Inches(left), Inches(top), text_width, text_height)
#     text_frame = textBox.text_frame
#     # center vertically
#     text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
#     # text_frame.text = text

#     # Style the text
#     paragraph = text_frame.paragraphs[0]
#     paragraph.alignment = PP_ALIGN.CENTER
#     run = paragraph.add_run()
#     run.text = text
#     run.font.size = Pt(36)
#     run.font.color.rgb = RGBColor(255, 255, 255)

# ppt.save('test.pptx')