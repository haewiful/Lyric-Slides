import tkinter as tk
from tkinter import messagebox
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

def create_presentation():
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
    while True:
        try:
            font_size = int(font_entry.get())
            break
        except ValueError:
            messagebox.showwarning("Input Error", "Font size has to be an integer.")
            return

    # get file name
    file_name = name_entry.get()
    if not file_name:
        messagebox.showwarning("Input Error", "Please give a file name.")
        return

    for idx in range(0, len(text_input), 2):
        text = text_input[idx]
        if idx+1 < len(text_input):
            text += "\n" + text_input[idx+1]
        create_slide(text, font_size)

    ppt.save(file_name+'.pptx')
    gui.destroy()

def create_slide(text, font_size):
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

    # Style the text
    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(255, 255, 255)


ppt = Presentation()

gui = tk.Tk()
gui.title("Slide Maker")

# Lyrics entry
lyrics_frame = tk.Frame(gui)
lyrics_frame.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="w")

tk.Label(lyrics_frame, text="Enter Lyrics").pack()
text_entry = tk.Text(lyrics_frame, width=50, height=20)
text_entry.pack(pady=5)

# font size
font_frame = tk.Frame(gui)
font_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="w")

font_default = 36
tk.Label(font_frame, text="Font Size:").pack(side="left")

font_entry = tk.Entry(font_frame, width=5)
font_entry.insert(0, str(font_default))
font_entry.pack(side="left", padx=5)

# file name
name_frame = tk.Frame(gui)
name_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="w")
tk.Label(name_frame, text="File Name:").pack(side="left")

name_entry = tk.Entry(name_frame, width=20)
name_entry.pack(side="left", padx=0)

tk.Label(name_frame, text=".pptx").pack(side="left", padx=0)

# Create presentation button
tk.Button(gui, text="Create Presentation", command=create_presentation).grid(row=3, columnspan=3, pady=20)

gui.mainloop()
