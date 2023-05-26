import customtkinter
from tkinter import *
from PIL import Image
import praw
from praw.models import MoreComments
import datetime
import pytz
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

red_json_path = resource_path('assets/devolved/devolved.json')
light_png_path = resource_path('assets/devolved/light.png')
dark_png_path = resource_path('assets/devolved/dark.png')
icon_path = resource_path('assets/devolved/icon.ico')
font_path = Path(__file__).parent / 'assets/font/CenzoFlareCond-Light.ttf'

reddit = praw.Reddit(
    client_id='CLIENT_ID',  # Replace with your actual client id
        client_secret='CLIENT_SECRET',  # Replace with your actual client secret
        user_agent='CLIENT_NAME'  # Input your client name
)

def retrieve_comments_stormont(submission_url, deadline_date):
    submission = reddit.submission(url=submission_url)

    total_comments = 0
    votes = {}

    deadline = datetime.datetime.strptime(deadline_date, "%d/%m/%Y %H:%M")
    deadline = pytz.timezone('Europe/London').localize(deadline)
    deadline = deadline.astimezone(pytz.UTC)
    deadline_timestamp = deadline.timestamp()

    for comment in submission.comments.list():
        if isinstance(comment, MoreComments):
            continue
        total_comments += 1
        if comment.author != "AutoModerator" and comment.created_utc <= deadline_timestamp:
            comment_body = comment.body.lower()
            if "no" in comment_body:
                votes[str(comment.author)] = "no"
            elif "aye" in comment_body:
                votes[str(comment.author)] = "aye"
                
                if "proxy for" in comment_body:
                    proxy_info = re.search(r"proxy for.*?/u/(\w+).*?=.*?(\w+)", comment_body)
                    if proxy_info:
                        proxied_username = proxy_info.group(1)
                        proxy_vote = proxy_info.group(2)
                        votes[proxied_username] = proxy_vote

            elif "abstain" in comment_body:
                votes[str(comment.author)] = "abstain"

                if "proxy for" in comment_body:
                    proxy_info = re.search(r"proxy for.*?/u/(\w+).*?=.*?(\w+)", comment_body)
                    if proxy_info:
                        proxied_username = proxy_info.group(1)
                        proxy_vote = proxy_info.group(2)
                        votes[proxied_username] = proxy_vote

    aye_votes = [username for username, vote in votes.items() if vote == "aye"]
    no_votes = [username for username, vote in votes.items() if vote == "no"]
    abstain_votes = [username for username, vote in votes.items() if vote == "abstain"]

    return total_comments, aye_votes, no_votes, abstain_votes

def retrieve_comments_senedd(submission_url, deadline_date):
    submission = reddit.submission(url=submission_url)

    total_comments = 0
    votes = {}

    deadline = datetime.datetime.strptime(deadline_date, "%d/%m/%Y %H:%M")
    deadline = pytz.timezone('Europe/London').localize(deadline)
    deadline = deadline.astimezone(pytz.UTC)
    deadline_timestamp = deadline.timestamp()

    for comment in submission.comments.list():
        if isinstance(comment, MoreComments):
            continue
        total_comments += 1
        if comment.author != "AutoModerator" and comment.created_utc <= deadline_timestamp:
            comment_body = comment.body.lower()
            if "against" in comment_body or "yn erbyn" in comment_body:
                votes[str(comment.author)] = "against"
            elif "for" in comment_body or "o blaid" in comment_body:
                votes[str(comment.author)] = "for"
                
                if "proxy for" in comment_body:
                    proxy_info = re.search(r"proxy for.*?/u/(\w+).*?=.*?(\w+)", comment_body)
                    if proxy_info:
                        proxied_username = proxy_info.group(1)
                        proxy_vote = proxy_info.group(2)
                        votes[proxied_username] = proxy_vote

            elif "abstain" in comment_body or "ymatal" in comment_body:
                votes[str(comment.author)] = "abstain"

                if "proxy for" in comment_body:
                    proxy_info = re.search(r"proxy for.*?/u/(\w+).*?=.*?(\w+)", comment_body)
                    if proxy_info:
                        proxied_username = proxy_info.group(1)
                        proxy_vote = proxy_info.group(2)
                        votes[proxied_username] = proxy_vote

    for_votes = [username for username, vote in votes.items() if vote == "for"]
    against_votes = [username for username, vote in votes.items() if vote == "against"]
    abstain_votes = [username for username, vote in votes.items() if vote == "abstain"]

    return total_comments, for_votes, against_votes, abstain_votes

def retrieve_comments_holyrood(submission_url, deadline_date):
    submission = reddit.submission(url=submission_url)

    total_comments = 0
    votes = {}

    deadline = datetime.datetime.strptime(deadline_date, "%d/%m/%Y %H:%M")
    deadline = pytz.timezone('Europe/London').localize(deadline)
    deadline = deadline.astimezone(pytz.UTC)
    deadline_timestamp = deadline.timestamp()

    for comment in submission.comments.list():
        if isinstance(comment, MoreComments):
            continue
        total_comments += 1
        if comment.author != "AutoModerator" and comment.created_utc <= deadline_timestamp:
            comment_body = comment.body.lower()
            if "against" in comment_body:
                votes[str(comment.author)] = "against"
            elif "for" in comment_body:
                votes[str(comment.author)] = "for"
                
                if "proxy for" in comment_body:
                    proxy_info = re.search(r"proxy for.*?/u/(\w+).*?=.*?(\w+)", comment_body)
                    if proxy_info:
                        proxied_username = proxy_info.group(1)
                        proxy_vote = proxy_info.group(2)
                        votes[proxied_username] = proxy_vote

            elif "abstain" in comment_body:
                votes[str(comment.author)] = "abstain"

                if "proxy for" in comment_body:
                    proxy_info = re.search(r"proxy for.*?/u/(\w+).*?=.*?(\w+)", comment_body)
                    if proxy_info:
                        proxied_username = proxy_info.group(1)
                        proxy_vote = proxy_info.group(2)
                        votes[proxied_username] = proxy_vote

    for_votes = [username for username, vote in votes.items() if vote == "for"]
    against_votes = [username for username, vote in votes.items() if vote == "against"]
    abstain_votes = [username for username, vote in votes.items() if vote == "abstain"]

    return total_comments, for_votes, against_votes, abstain_votes

def on_button_click_stormont():
    submission_url = entry1.get()
    deadline_date = entry2.get()

    if not submission_url or not deadline_date:
        result_text = "Error: Please enter both the Division URL and Division End Date and Time."
    else:
        if "reddit.com/r/MStormontVote/" not in submission_url:
            result_text = "Invalid URL. Please enter a URL from /r/MStormontVote."
            result.delete('1.0', END)
            result.insert(INSERT, result_text)
            return
        
        total_comments, aye_votes, no_votes, abstain_votes = retrieve_comments_stormont(submission_url, deadline_date)

        result_text = f"TOTAL VOTES: {total_comments}\n"
        result_text += f"TOTAL AYES: {len(aye_votes)}\n"
        result_text += f"TOTAL NOES: {len(no_votes)}\n"
        result_text += f"TOTAL ABSTAINS: {len(abstain_votes)}\n\n"
        result_text += "AYE {NI}:\n"
        result_text += "\n".join(aye_votes)
        result_text += "\n\nNO {NI}:\n"
        result_text += "\n".join(no_votes)
        result_text += "\n\nABSTAIN {NI}:\n"
        result_text += "\n".join(abstain_votes)

    result.delete('1.0', END)
    result.insert(INSERT, result_text)

def on_button_click_senedd():
    submission_url = entry1.get()
    deadline_date = entry2.get()

    if not submission_url or not deadline_date:
        result_text = "Error: Please enter both the Division URL and Division End Date and Time."
    else:
        if "reddit.com/r/MHOCSeneddVote/" not in submission_url:
            result_text = "Invalid URL. Please enter a URL from /r/MHOCSeneddVote."
            result.delete('1.0', END)
            result.insert(INSERT, result_text)
            return
        
        total_comments, for_votes, against_votes, abstain_votes = retrieve_comments_senedd(submission_url, deadline_date)

        result_text = f"TOTAL VOTES: {total_comments}\n"
        result_text += f"TOTAL FORS: {len(for_votes)}\n"
        result_text += f"TOTAL AGAINSTS: {len(against_votes)}\n"
        result_text += f"TOTAL ABSTAINS: {len(abstain_votes)}\n\n"
        result_text += "FOR {CYM}:\n"
        result_text += "\n".join(for_votes)
        result_text += "\n\nAGAINST {CYM}:\n"
        result_text += "\n".join(against_votes)
        result_text += "\n\nABSTAIN {CYM}:\n"
        result_text += "\n".join(abstain_votes)

    result.delete('1.0', END)
    result.insert(INSERT, result_text)

def on_button_click_holyrood():
    submission_url = entry1.get()
    deadline_date = entry2.get()

    if not submission_url or not deadline_date:
        result_text = "Error: Please enter both the Division URL and Division End Date and Time."
    else:
        if "reddit.com/r/MHOCHolyroodVote/" not in submission_url:
            result_text = "Invalid URL. Please enter a URL from /r/MHOCHolyroodVote."
            result.delete('1.0', END)
            result.insert(INSERT, result_text)
            return
        
        total_comments, for_votes, against_votes, abstain_votes = retrieve_comments_holyrood(submission_url, deadline_date)

        result_text = f"TOTAL VOTES: {total_comments}\n"
        result_text += f"TOTAL FORS: {len(for_votes)}\n"
        result_text += f"TOTAL AGAINSTS: {len(against_votes)}\n"
        result_text += f"TOTAL ABSTAINS: {len(abstain_votes)}\n\n"
        result_text += "FOR:\n"
        result_text += "\n".join(for_votes)
        result_text += "\n\nAGAINST:\n"
        result_text += "\n".join(against_votes)
        result_text += "\n\nABSTAIN:\n"
        result_text += "\n".join(abstain_votes)

    result.delete('1.0', END)
    result.insert(INSERT, result_text)

app = customtkinter.CTk()
app.title("VoteVerdict • Parliamentary Toolkit • Devolved Edition")
app.geometry("550x810")
app.iconbitmap(default=icon_path)
app.resizable(False, False)
app.grid_columnconfigure(0, weight=1)

customtkinter.set_appearance_mode("system")
customtkinter.set_appearance_mode("dark")
customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme("assets/theme/devolved.json")

my_image = customtkinter.CTkImage(light_image=Image.open(resource_path("assets/devolved/light.png")),
                                  dark_image=Image.open(resource_path("assets/devolved/dark.png")),
                                  size=(300, 133))

image_label = customtkinter.CTkLabel(app, image=my_image, text="")
image_label.grid(row=1, column=0, padx=20, pady=(0, 10))

label = customtkinter.CTkLabel(app, font=("Cenzo Flare Cond Light", 18), text="Division URL:", fg_color="transparent")
label.grid(row=2, column=0, padx=20, pady=(10, 0))
entry1 = customtkinter.CTkEntry(app, font=("Cenzo Flare Cond Light", 15), placeholder_text="ENTER URL HERE")
entry1.grid(row=3, column=0, padx=20, pady=0)

label = customtkinter.CTkLabel(app, font=("Cenzo Flare Cond Light", 18), text="Division End Date and Time:", fg_color="transparent")
label.grid(row=4, column=0, padx=20, pady=(10, 0))
entry2 = customtkinter.CTkEntry(app, font=("Cenzo Flare Cond Light", 15), placeholder_text="DD/MM/YYYY HH:MM")
entry2.grid(row=5, column=0, padx=20, pady=0)

frame = customtkinter.CTkFrame(app, fg_color="transparent")
frame.grid(row=6, column=0, columnspan=2, padx=0, pady=0)

button = customtkinter.CTkButton(frame, font=("Cenzo Flare Cond Light", 21), fg_color="#5485C0", hover_color="#436A99", width=10, text="stormont", command=on_button_click_stormont)
button.grid(row=0, column=0, padx=2.5, pady=20)

button = customtkinter.CTkButton(frame, font=("Cenzo Flare Cond Light", 21), fg_color="#481E6F", hover_color="#32154D", width=10, text="holyrood", command=on_button_click_holyrood)
button.grid(row=0, column=1, padx=2.5, pady=20)

button = customtkinter.CTkButton(frame, font=("Cenzo Flare Cond Light", 21), fg_color="#820045", hover_color="#5B0030", width=10, text="senedd", command=on_button_click_senedd)
button.grid(row=0, column=2, padx=2.5, pady=20)

result = customtkinter.CTkTextbox(app, font=("Cenzo Flare Cond Light", 15), height=330, width=450)
result.grid(row=7, column=0, columnspan=2, padx=20, pady=0)

appearance_mode_button = customtkinter.CTkSegmentedButton(app, font=("Cenzo Flare Cond Light", 21), values=["light", "dark"], command=lambda v: customtkinter.set_appearance_mode(v))
appearance_mode_button.grid(row=8, column=0, columnspan=4, padx=20, pady=20)

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