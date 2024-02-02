import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys
from PIL import Image, ImageTk
import os
import threading
import pyaudio
from vosk import Model, KaldiRecognizer
import pyautogui
import json
import time
import inflect
import re

# Create an inflect engine
p = inflect.engine()

def number_to_words(text):
    """
    Convert numeric text to words. E.g., '1' to 'one'.
    """
    words = []
    for word in text.split():
        if word.isdigit():
            words.append(p.number_to_words(word))
        else:
            words.append(word)
    return ' '.join(words)


def open_powerpoint():
    # Open a dialog to choose a PowerPoint file
    powerpoint_file_path = filedialog.askopenfilename(
        title="Select PowerPoint file",
        filetypes=[("PowerPoint presentations", "*.pptx;*.ppt")])

    # Check if a file was selected
    if powerpoint_file_path:
        try:
            # Open the PowerPoint file with the default application
            if os.name == 'nt':  # Windows
                os.startfile(powerpoint_file_path)
                # Wait for the PowerPoint application to open the file
                time.sleep(3)  # Adjust the delay as needed
                # Simulate pressing F5 to start the slide show
                pyautogui.press('f5')
            elif os.name == 'posix':  # macOS or Linux
                opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
                subprocess.run([opener, powerpoint_file_path])
                # On macOS, you might need to use a different key sequence
                # time.sleep(3)  # Adjust the delay as needed
                # pyautogui.hotkey('ctrl', 'cmd', 's')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open the file: {e}")

def voice_control():
    # Set up Vosk with the model
    model = Model("model")
    recognizer = KaldiRecognizer(model, 16000)
    p = pyaudio.PyAudio()
    stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=8192)
    stream.start_stream()

    # Continuous recognition loop
    while not exit_thread:  # exit_thread is a global flag to stop the thread
        data = stream.read(4096, exception_on_overflow=False)
        if recognizer.AcceptWaveform(data):
            result = recognizer.Result()
            result_dict = json.loads(result)
            text = result_dict.get('text', '')
            process_speech(text)

    # Clean up
    stream.stop_stream()
    stream.close()
    p.terminate()

# Add global variables to store the custom commands
custom_command_left = "back"  # default command for left
custom_command_right = "next"  # default command for right

# ... your existing functions ...


def word_to_num(word):
    """
    Convert a word representation of a number to its numeric form.
    Currently supports numbers from one to ten.
    """
    mapping = {
        "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
        "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10
    }
    return mapping.get(word.lower())

def process_speech(text):
    print(f"Recognized: {text}")  # Print all recognized text for debugging
    global custom_command_left, custom_command_right
    
    # Existing commands for next and previous
    if custom_command_right in text:
        pyautogui.press('right')
        print(f"Command: {custom_command_right}")  # Print command execution confirmation
    elif custom_command_left in text:
        pyautogui.press('left')
        print(f"Command: {custom_command_left}")  # Print command execution confirmation

    # Convert numbers in text to words
    text_in_words = number_to_words(text)

    # Check if the speech contains a slide number command
    slide_number_match = re.search(r'number (one|two|three|four|five|six|seven|eight|nine|ten|\d+)', text_in_words)
    if slide_number_match:
        slide_number_word = slide_number_match.group(1)

        # Convert word to number if necessary
        if slide_number_word.isdigit():
            slide_number = int(slide_number_word)
        else:
            slide_number = word_to_num(slide_number_word)

        if slide_number is not None:
            jump_to_slide(slide_number)
            print(f"Jumping to slide: {slide_number}")
        return
    
def jump_to_slide(slide_number):
    """
    Function to jump to a specific slide number.
    """
    pyautogui.typewrite(str(slide_number))  # Type the slide number
    pyautogui.press('enter')  # Press Enter to confirm

# Add GUI elements for custom command input
def save_custom_commands():
    global custom_command_left, custom_command_right
    custom_command_left = left_command_entry.get()
    custom_command_right = right_command_entry.get()
    messagebox.showinfo("Info", "Commands updated successfully")

exit_thread = False

def on_closing():
    global exit_thread
    exit_thread = True
    root.destroy()

root = tk.Tk()
root.title("VoicePoint")

# Improved window dimensions and allowing resizing
root.geometry("500x300")
root.resizable(True, True)

# Use a more appealing color scheme
background_color = "#333333"  # Dark background
text_color = "#ffffff"  # White text
button_color = "#4CAF50"  # Green buttons
entry_background = "#ffffff"  # White entry fields
root.configure(bg=background_color)

# Use custom fonts
title_font = ("Helvetica", 20, "bold")
label_font = ("Helvetica", 12)
button_font = ("Helvetica", 12, "bold")

# Create a label with a custom font
title_label = tk.Label(root, text="Welcome to VoicePoint", font=title_font, fg=text_color, bg=background_color)
title_label.pack(pady=20)

# Frame for command input
command_frame = tk.Frame(root, bg=background_color)
command_frame.pack(pady=10)

# Left command widgets
left_command_label = tk.Label(command_frame, text="Command for LEFT", font=label_font, fg=text_color, bg=background_color)
left_command_label.grid(row=0, column=0, padx=10, pady=10)
left_command_entry = tk.Entry(command_frame, bg=entry_background)
left_command_entry.grid(row=0, column=1, padx=10, pady=10)

# Right command widgets
right_command_label = tk.Label(command_frame, text="Command for RIGHT", font=label_font, fg=text_color, bg=background_color)
right_command_label.grid(row=1, column=0, padx=10, pady=10)
right_command_entry = tk.Entry(command_frame, bg=entry_background)
right_command_entry.grid(row=1, column=1, padx=10, pady=10)

# Button to open PowerPoint
open_button = tk.Button(root, text="Open PowerPoint", command=open_powerpoint, bg=button_color, font=button_font)
open_button.pack(pady=10)

# Button to save commands
save_commands_button = tk.Button(root, text="Save Commands", command=save_custom_commands, bg=button_color, font=button_font)
save_commands_button.pack(pady=10)





# Start the voice control thread
voice_thread = threading.Thread(target=voice_control)
voice_thread.start()

# Set a callback for when the window is closed
root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()

# Wait for the voice control thread to finish before exiting the program
voice_thread.join()
