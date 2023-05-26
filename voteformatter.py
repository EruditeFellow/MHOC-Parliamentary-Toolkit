import customtkinter
from tkinter import *
from PIL import Image
import re
import pyglet
from pathlib import Path
import sys
import os

pyglet.options['win32_gdi_font'] = True
font_path = Path(__file__).parent / 'assets/font/CenzoFlareCond-Light.ttf'
pyglet.font.add_file(str(font_path))

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

red_json_path = resource_path('assets/theme/serenity.json')
light_png_path = resource_path('assets/formatter/light.png')
dark_png_path = resource_path('assets/formatter/dark.png')
icon_path = resource_path('assets/formatter/icon.ico')
font_path = Path(__file__).parent / 'assets/font/CenzoFlareCond-Light.ttf'

def process_text(mode):
    input_text = input.get("1.0", "end-1c")

    if not input_text.strip():
        error_message = "Error: Please input voting record before choosing a mode."
        output.delete("1.0", "end")
        output.insert("1.0", error_message)
        return

    if mode == "mhol":
        if not re.search(r'CONTENT:(.*?)NOT CONTENT:', input_text, re.DOTALL) or \
                not re.search(r'NOT CONTENT:(.*?)PRESENT:', input_text, re.DOTALL):
            error_message = "Invalid input: MHoL mode cannot be used with this input type."
            output.delete("1.0", "end")
            output.insert("1.0", error_message)
            return

        contents = re.findall(r'CONTENT:(.*?)NOT CONTENT:', input_text, re.DOTALL)[0].strip().split('\n')
        contents = [f"'{name.strip()}'" for name in contents if name.strip() != '']

        not_contents = re.findall(r'NOT CONTENT:(.*?)PRESENT:', input_text, re.DOTALL)[0].strip().split('\n')
        not_contents = [f"'{name.strip()}'" for name in not_contents if name.strip() != '']

        presents = re.findall(r'PRESENT:(.*?)$', input_text, re.DOTALL)[0].strip().split('\n')
        presents = [f"'{name.strip()}'" for name in presents if name.strip() != '']

        output_text = f"""function addVotes() {{
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getSheetByName('L: 19th Term Voting Record'); // Enter sheet name
      
      var contentVotes = [{', '.join(contents)}].map(name => name.toLowerCase());

      var notContentVotes = [{', '.join(not_contents)}].map(name => name.toLowerCase());

      var presentVotes = [{', '.join(presents)}].map(name => name.toLowerCase());

      var range = sheet.getRange('C4:C46'); // Adjusted range from C4 to C46 - this row is for usernames
      var values = range.getValues();

      for (var i = 0; i < values.length; i++) {{
        var name = values[i][0].toLowerCase();
        var voteCell = sheet.getRange('AI' + (i + 4)); // Adjusted to start from AI(4) - Change it depending on the column you want to fill out
        if (contentVotes.includes(name)) {{
          voteCell.setValue('CON');
        }} else if (notContentVotes.includes(name)) {{
          voteCell.setValue('NOT');
        }} else if (presentVotes.includes(name)) {{
          voteCell.setValue('PRE');
        }} else {{
          voteCell.setValue('DNV'); // Clear the cell if the name does not match any vote
        }}
      }}
    }}"""

    elif mode == "mhoc":
        if not re.search(r'AYE:(.*?)NO:', input_text, re.DOTALL) or \
                not re.search(r'NO:(.*?)ABSTAIN:', input_text, re.DOTALL) or \
                not re.search(r'ABSTAIN:(.*?)$', input_text, re.DOTALL):
            error_message = "Invalid input: MHoC mode cannot be used with this input type."
            output.delete("1.0", "end")
            output.insert("1.0", error_message)
            return

        ayes = re.findall(r'AYE:(.*?)NO:', input_text, re.DOTALL)[0].strip().split('\n')
        ayes = [f"'{name.strip()}'" for name in ayes if name.strip() != '']

        noes = re.findall(r'NO:(.*?)ABSTAIN:', input_text, re.DOTALL)[0].strip().split('\n')
        noes = [f"'{name.strip()}'" for name in noes if name.strip() != '']

        abstains = re.findall(r'ABSTAIN:(.*?)$', input_text, re.DOTALL)[0].strip().split('\n')
        abstains = [f"'{name.strip()}'" for name in abstains if name.strip() != '']
    
        output_text = f"""function addVotes() {{
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getSheetByName('C: 19th Term Voting Record'); // Enter sheet name

      var ayeVotes = [{','.join(ayes)}].map(name => name.toLowerCase());

      var noVotes = [{','.join(noes)}].map(name => name.toLowerCase());

      var abstainVotes = [{','.join(abstains)}].map(name => name.toLowerCase());

      var range = sheet.getRange('C4:C153'); // Adjusted range from C4 to C153 - this row is for usernames
      var values = range.getValues();
      var styles = range.getTextStyles(); // Get the text styles in the range

      var lastVote = 'DNV';
      var lastName = '';

      for (var i = 0; i < values.length; i++) {{
        var name = values[i][0].toLowerCase();
        var voteCell = sheet.getRange('T' + (i + 4)); // Adjusted to start from T(4) - Change it depending on the column you want to fill out

        if (name === '') {{
          name = lastName;
        }} else {{
          lastName = name;
        }}

        if (styles[i][0].isStrikethrough()) {{ // Check if username has a strikethrough
          voteCell.setValue('N/A');
        }} else {{
          if (ayeVotes.includes(name)) {{
            voteCell.setValue('AYE');
            lastVote = 'AYE';
          }} else if (noVotes.includes(name)) {{
            voteCell.setValue('NO');
            lastVote = 'NO';
          }} else if (abstainVotes.includes(name)) {{
            voteCell.setValue('ABS');
            lastVote = 'ABS';
          }} else {{
            voteCell.setValue('DNV'); 
            lastVote = 'DNV';
          }}
        }}
      }}
    }}"""
    
    elif mode == "stormont":
        if not re.search(r'AYE {NI}:(.*?)NO {NI}:', input_text, re.DOTALL) or \
                not re.search(r'NO {NI}:(.*?)ABSTAIN {NI}:', input_text, re.DOTALL) or \
                not re.search(r'ABSTAIN {NI}:(.*?)$', input_text, re.DOTALL):
            error_message = "Invalid input: Stormont mode cannot be used with this input type."
            output.delete("1.0", "end")
            output.insert("1.0", error_message)
            return

        ayes = re.findall(r'AYE {NI}:(.*?)NO {NI}:', input_text, re.DOTALL)[0].strip().split('\n')
        ayes = [f"'{name.strip()}'" for name in ayes if name.strip() != '']

        noes = re.findall(r'NO {NI}:(.*?)ABSTAIN {NI}:', input_text, re.DOTALL)[0].strip().split('\n')
        noes = [f"'{name.strip()}'" for name in noes if name.strip() != '']

        abstains = re.findall(r'ABSTAIN {NI}:(.*?)$', input_text, re.DOTALL)[0].strip().split('\n')
        abstains = [f"'{name.strip()}'" for name in abstains if name.strip() != '']
    
        output_text = f"""function addVotes() {{
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getSheetByName('13th Assembly Voting'); // Enter sheet name

      var ayeVotes = [{','.join(ayes)}].map(name => name.toLowerCase());

      var noVotes = [{','.join(noes)}].map(name => name.toLowerCase());

      var abstainVotes = [{','.join(abstains)}].map(name => name.toLowerCase());

      var range = sheet.getRange('C5:C28'); // Adjusted range from C5 to C28 - this row is for usernames
      var values = range.getValues();
      var styles = range.getTextStyles(); // Get the text styles in the range

      var lastVote = 'DNV';
      var lastName = '';

      for (var i = 0; i < values.length; i++) {{
        var name = values[i][0].toLowerCase();
        var voteCell = sheet.getRange('AT' + (i + 5)); // Adjusted to start from AT(5) - Change it depending on the column you want to fill out

        if (name === '') {{
          name = lastName;
        }} else {{
          lastName = name;
        }}

        if (styles[i][0].isStrikethrough()) {{ // Check if username has a strikethrough
          voteCell.setValue('N/A');
        }} else {{
          if (ayeVotes.includes(name)) {{
            voteCell.setValue('AYE');
            lastVote = 'AYE';
          }} else if (noVotes.includes(name)) {{
            voteCell.setValue('NO');
            lastVote = 'NO';
          }} else if (abstainVotes.includes(name)) {{
            voteCell.setValue('ABS');
            lastVote = 'ABS';
          }} else {{
            voteCell.setValue('DNV'); 
            lastVote = 'DNV';
          }}
        }}
      }}
    }}"""

    elif mode == "holyrood":
        if not re.search(r'FOR:(.*?)AGAINST:', input_text, re.DOTALL) or \
                not re.search(r'AGAINST:(.*?)ABSTAIN:', input_text, re.DOTALL) or \
                not re.search(r'ABSTAIN:(.*?)$', input_text, re.DOTALL):
            error_message = "Invalid input: Holyrood mode cannot be used with this input type."
            output.delete("1.0", "end")
            output.insert("1.0", error_message)
            return

        fors = re.findall(r'FOR:(.*?)AGAINST:', input_text, re.DOTALL)[0].strip().split('\n')
        fors = [f"'{name.strip()}'" for name in fors if name.strip() != '']

        againsts = re.findall(r'AGAINST:(.*?)ABSTAIN:', input_text, re.DOTALL)[0].strip().split('\n')
        againsts = [f"'{name.strip()}'" for name in againsts if name.strip() != '']

        abstains = re.findall(r'ABSTAIN:(.*?)$', input_text, re.DOTALL)[0].strip().split('\n')
        abstains = [f"'{name.strip()}'" for name in abstains if name.strip() != '']
    
        output_text = f"""function addVotes() {{
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getSheetByName('12th Parliament Voting Record'); // Enter sheet name

      var forVotes = [{','.join(fors)}].map(name => name.toLowerCase());

      var againstVotes = [{','.join(againsts)}].map(name => name.toLowerCase());

      var abstainVotes = [{','.join(abstains)}].map(name => name.toLowerCase());

      var range = sheet.getRange('C4:C38'); // Adjusted range from C4 to C38 - this row is for usernames
      var values = range.getValues();
      var styles = range.getTextStyles(); // Get the text styles in the range

      var lastVote = 'DNV';
      var lastName = '';

      for (var i = 0; i < values.length; i++) {{
        var name = values[i][0].toLowerCase();
        var voteCell = sheet.getRange('BD' + (i + 4)); // Adjusted to start from BD(4) - Change it depending on the column you want to fill out

        if (name === '') {{
          name = lastName;
        }} else {{
          lastName = name;
        }}

        if (styles[i][0].isStrikethrough()) {{ // Check if username has a strikethrough
          voteCell.setValue('N/A');
        }} else {{
          if (forVotes.includes(name)) {{
            voteCell.setValue('For');
            lastVote = 'For';
          }} else if (againstVotes.includes(name)) {{
            voteCell.setValue('Against');
            lastVote = 'Against';
          }} else if (abstainVotes.includes(name)) {{
            voteCell.setValue('Abstain');
            lastVote = 'Abstain';
          }} else {{
            voteCell.setValue('DNV'); 
            lastVote = 'DNV';
          }}
        }}
      }}
    }}"""
        
    elif mode == "senedd":
        if not re.search(r'FOR {CYM}:(.*?)AGAINST {CYM}:', input_text, re.DOTALL) or \
                not re.search(r'AGAINST {CYM}:(.*?)ABSTAIN {CYM}:', input_text, re.DOTALL) or \
                not re.search(r'ABSTAIN {CYM}:(.*?)$', input_text, re.DOTALL):
            error_message = "Invalid input: Senedd mode cannot be used with this input type."
            output.delete("1.0", "end")
            output.insert("1.0", error_message)
            return

        fors = re.findall(r'FOR {CYM}:(.*?)AGAINST {CYM}:', input_text, re.DOTALL)[0].strip().split('\n')
        fors = [f"'{name.strip()}'" for name in fors if name.strip() != '']

        againsts = re.findall(r'AGAINST {CYM}:(.*?)ABSTAIN {CYM}:', input_text, re.DOTALL)[0].strip().split('\n')
        againsts = [f"'{name.strip()}'" for name in againsts if name.strip() != '']

        abstains = re.findall(r'ABSTAIN {CYM}:(.*?)$', input_text, re.DOTALL)[0].strip().split('\n')
        abstains = [f"'{name.strip()}'" for name in abstains if name.strip() != '']
    
        output_text = f"""function addVotes() {{
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getSheetByName('Voting Record of the 9th Senedd'); // Enter sheet name

      var forVotes = [{','.join(fors)}].map(name => name.toLowerCase());

      var againstVotes = [{','.join(againsts)}].map(name => name.toLowerCase());

      var abstainVotes = [{','.join(abstains)}].map(name => name.toLowerCase());

      var range = sheet.getRange('D4:D30'); // Adjusted range from D4 to D30 - this row is for usernames
      var values = range.getValues();
      var styles = range.getTextStyles(); // Get the text styles in the range

      var lastVote = 'DNV';
      var lastName = '';

      for (var i = 0; i < values.length; i++) {{
        var name = values[i][0].toLowerCase();
        var voteCell = sheet.getRange('AX' + (i + 4)); // Adjusted to start from AX(4) - Change it depending on the column you want to fill out

        if (name === '') {{
          name = lastName;
        }} else {{
          lastName = name;
        }}

        if (styles[i][0].isStrikethrough()) {{ // Check if username has a strikethrough
          voteCell.setValue('N/A');
        }} else {{
          if (forVotes.includes(name)) {{
            voteCell.setValue('For');
            lastVote = 'For';
          }} else if (againstVotes.includes(name)) {{
            voteCell.setValue('Against');
            lastVote = 'Against';
          }} else if (abstainVotes.includes(name)) {{
            voteCell.setValue('Abstain');
            lastVote = 'Abstain';
          }} else {{
            voteCell.setValue('DNV'); 
            lastVote = 'DNV';
          }}
        }}
      }}
    }}"""
    else:
        error_message = "Invalid mode"
        output.delete("1.0", "end")
        output.insert("1.0", error_message)
        return

    output.delete("1.0", "end")
    output.insert("1.0", output_text)

app = customtkinter.CTk()
app.title("Vote Formatter • Parliamentary Toolkit")
app.geometry("550x810")
app.iconbitmap(default=icon_path)
app.resizable(False, False)
app.grid_columnconfigure(0, weight=1)

customtkinter.set_appearance_mode("system")
customtkinter.set_appearance_mode("dark")
customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme("assets/theme/serenity.json")

my_image = customtkinter.CTkImage(light_image=Image.open(resource_path("assets/formatter/light.png")),
                                  dark_image=Image.open(resource_path("assets/formatter/dark.png")),
                                  size=(300, 133))

label = customtkinter.CTkLabel(app, text="", fg_color="transparent")
label.grid(row=0, column=0, padx=0, pady=(0, 0))

image_label = customtkinter.CTkLabel(app, image=my_image, text="")
image_label.grid(row=1, column=0, padx=20, pady=(0, 10))

label = customtkinter.CTkLabel(app, font=("Cenzo Flare Cond Light", 18), text="Input Voting Record:", fg_color="transparent")
label.grid(row=2, column=0, padx=20, pady=(10, 0))
input = customtkinter.CTkTextbox(app, font=("Cenzo Flare Cond Light", 15), height=220, width=450)
input.grid(row=3, column=0, columnspan=2, padx=20, pady=0)

# button = customtkinter.CTkButton(app, font=("Cenzo Flare Cond Light", 21), text="generate output", command=process_text)
# button.grid(row=4, column=0, padx=20, pady=20)

label = customtkinter.CTkLabel(app, font=("Cenzo Flare Cond Light", 18), text="Output:", fg_color="transparent")
label.grid(row=5, column=0, padx=20, pady=(10, 0))
output = customtkinter.CTkTextbox(app, font=("Cenzo Flare Cond Light", 15), height=220, width=450)
output.grid(row=5, column=0, columnspan=2, padx=20, pady=0)

appearance_mode_button = customtkinter.CTkSegmentedButton(app, font=("Cenzo Flare Cond Light", 21), values=["light", "dark"], command=lambda v: customtkinter.set_appearance_mode(v))
appearance_mode_button.grid(row=8, column=0, columnspan=4, padx=20, pady=20)

frame = customtkinter.CTkFrame(app, fg_color="transparent")
frame.grid(row=4, column=0, columnspan=2, padx=0, pady=0)

switch_button = customtkinter.CTkButton(frame, font=("Cenzo Flare Cond Light", 21), fg_color="#BC5154", hover_color="#9B4043", width=10, text="mhol", command=lambda: process_text("mhol"))
switch_button.grid(row=0, column=0, padx=2.5, pady=20)

switch_button = customtkinter.CTkButton(frame, font=("Cenzo Flare Cond Light", 21), fg_color="#748B74", hover_color="#5D6E5D", width=10, text="mhoc", command=lambda: process_text("mhoc"))
switch_button.grid(row=0, column=1, padx=2.5, pady=20)

switch_button = customtkinter.CTkButton(frame, font=("Cenzo Flare Cond Light", 21), fg_color="#5485C0", hover_color="#436A99", width=10, text="stormont", command=lambda: process_text("stormont"))
switch_button.grid(row=0, column=2, padx=2.5, pady=20)

switch_button = customtkinter.CTkButton(frame, font=("Cenzo Flare Cond Light", 21), fg_color="#481E6F", hover_color="#32154D", width=10, text="holyrood", command=lambda: process_text("holyrood"))
switch_button.grid(row=0, column=3, padx=2.5, pady=20)

switch_button = customtkinter.CTkButton(frame, font=("Cenzo Flare Cond Light", 21), fg_color="#820045", hover_color="#5B0030", width=10, text="senedd", command=lambda: process_text("senedd"))
switch_button.grid(row=0, column=4, padx=2.5, pady=20)

frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)

screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()

window_width = 550
window_height = 810

x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 3

app.geometry(f"{window_width}x{window_height}+{x}+{y}")

app.grid_rowconfigure(0, weight=1)
app.grid_rowconfigure(8, weight=1)

app.mainloop()