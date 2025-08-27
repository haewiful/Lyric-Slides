import tkinter as tk
from tkinter import messagebox, filedialog, font, ttk
# from tkinter import filedialog
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
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
    
    # # get file name
    # file_name = name_entry.get()
    # if not file_name:
    #     messagebox.showwarning("Input Error", "Please give a file name.")
    #     return

    # get font family
    font_family = font_family_combobox.get()

    # textbox width
    text_box_width = ppt.slide_width.inches
    # Approximate width per character (this is a rough estimate)
    avg_char_width = font_size * 0.75  # Average width of characters (in points)
    # Convert text box width from inches to points (1 inch = 72 points)
    text_box_width_in_points = text_box_width * 72
    # Estimate the number of characters per line (for korean)
    chars_per_line = int(text_box_width_in_points / avg_char_width)

    # for idx in range(0, len(text_input), 1):
    idx = 0
    while(idx < len(text_input)):
        text = text_input[idx]
        if idx+1 >= len(text_input):
            break
 
        # print("len(text):", len(text))
        # print("chars per line:", chars_per_line)
        # text doesn't overflow & next line doesn't overflow
        if len(text) < chars_per_line and len(text_input[idx+1]) < chars_per_line:
           text += "\n" + text_input[idx+1]
           idx+=1
        print(text)
        create_slide(text, font_size, font_family)
        idx+=1
    
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pptx",
        filetypes=[("PowerPoint files", "*.pptx")],
        title="Save Presentation As"
    )

    if not save_path:
        return
        
    # ppt.save('./'+file_name+'.pptx')
    ppt.save(save_path)
    gui.destroy()

def create_slide(text, font_size, font_family):
    slide_layout = ppt.slide_layouts[6]
    slide = ppt.slides.add_slide(slide_layout)

    # slide dimensions
    slide_width = ppt.slide_width.inches
    slide_height = ppt.slide_height.inches

    # text box dimensions
    # text_width = Inches(slide_width)
    text_height = Cm(4.37)
    
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
    run.font.name = font_family
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(255, 255, 255)


ppt = Presentation()

# Set the slide size (example: 16:9 widescreen using Inches)
ppt.slide_width = Inches(13.33)  # Width
ppt.slide_height = Inches(7.5)  # Height

text_width = Inches(ppt.slide_width.inches)

gui = tk.Tk()
gui.title("Slide Maker")

# Lyrics entry
lyrics_frame = tk.Frame(gui)
lyrics_frame.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="w")

tk.Label(lyrics_frame, text="Enter Lyrics").pack()
text_entry = tk.Text(lyrics_frame, width=50, height=20)
text_entry.pack(pady=5)

# font size
font_size_frame = tk.Frame(gui)
font_size_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="w")

font_default = 48
tk.Label(font_size_frame, text="Font Size:").pack(side="left")

font_entry = tk.Entry(font_size_frame, width=5)
font_entry.insert(0, str(font_default))
font_entry.pack(side="left", padx=5)

# font family selection
font_family_frame = tk.Frame(gui)
font_family_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="w")

tk.Label(font_family_frame, text="Font:").pack(side="left")

# Dynamically get the list of fonts from the system
font_families = sorted(list(font.families()))

font_family_combobox = ttk.Combobox(font_family_frame, values=font_families, state="readonly")
# Check if a common font is in the list to set as a default
if "Arial" in font_families:
    font_family_combobox.set("Arial")
else:
    font_family_combobox.set(font_families[0])

font_family_combobox.pack(side="left", padx=5)

# # file name
# name_frame = tk.Frame(gui)
# name_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="w")
# tk.Label(name_frame, text="File Name:").pack(side="left")

# name_entry = tk.Entry(name_frame, width=20)
# name_entry.pack(side="left", padx=0)

# tk.Label(name_frame, text=".pptx").pack(side="left", padx=0)

# Create presentation button
tk.Button(gui, text="Create Presentation", command=create_presentation).grid(row=3, columnspan=3, pady=20)

gui.mainloop()
