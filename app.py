import tkinter
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import tkinter.font as tkFont
import sv_ttk
import uuid
import os
import win32com.client
import asyncio
import edge_tts
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip


def select_output_directory():
    selected_directory = filedialog.askdirectory()
    if selected_directory:
        output_dir_entry.delete(0, tkinter.END)
        output_dir_entry.insert(0, selected_directory)


def select_input_ppt():
    selected_file = filedialog.askopenfilename(
        filetypes=[("PowerPoint files", "*.ppt *.pptx")]
    )
    if selected_file:
        ppt_entry.delete(0, tkinter.END)
        ppt_entry.insert(0, selected_file)


def on_submit():
    ppt_path = ppt_entry.get()
    ai_voice = voice_combobox.get()
    output_dir = output_dir_entry.get()
    video_quality = video_quality_combobox.get()

    if ppt_path and ai_voice and output_dir and video_quality:
        try:
            unique_id = str(uuid.uuid4())
            ppt_file_name = os.path.splitext(os.path.basename(ppt_path))[0]
            output_video_file_path = f"{output_dir}/{ppt_file_name}.mp4"

            os.makedirs(os.path.abspath(output_dir), exist_ok=True)
            os.makedirs(os.path.abspath(output_dir) + "/temp", exist_ok=True)

            match video_quality:
                case "720p":
                    act_video_quality = (1280, 720)
                case "480p":
                    act_video_quality = (720, 480)
                case _:
                    act_video_quality = (1920, 1080)

            slide_data = extract_slide_images(os.path.abspath(ppt_path),
                                              os.path.abspath(output_dir),
                                              unique_id,
                                              ppt_file_name)

            generate_video_from_slides(slide_data,
                                       os.path.abspath(output_video_file_path),
                                       tts_enabled=True,
                                       video_resolution=act_video_quality,
                                       ai_voice=ai_voice)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
    else:
        messagebox.showerror("Error", f"All the fields are mandatory. Make sure all the fields are filled")


async def tts(text, filename, ai_voice="en-US-AvaNeural"):
    communicate = edge_tts.Communicate(text, ai_voice)
    await communicate.save(filename)
    write_to_logs(f"-- Successfully created AI narration for {filename}\n")


def extract_slide_images(ppt_path, output_folder, unique_id, filename, image_width=1920, image_height=1080):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1

    presentation = powerpoint.Presentations.Open(ppt_path)

    slide_data = []

    for i, slide in enumerate(presentation.Slides):
        # Save slide as image
        image_path = os.path.join(output_folder, f"temp/{unique_id}_{filename}_slide_{i + 1}.png")
        slide.Export(image_path, "PNG", image_width, image_height)

        slide_text = ""
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                slide_text += shape.TextFrame.TextRange.Text + " "
        write_to_logs(f"-- Successfully converted Slide {i+1} to Image\n")
        slide_data.append((image_path, slide_text.strip()))

    presentation.Close()
    powerpoint.Quit()

    return slide_data


def generate_video_from_slides(slide_data, output_video_path, tts_enabled=False, video_resolution=(1920, 1080), ai_voice="en-US-AvaNeural"):
    clips = []
    for image_path, slide_text in slide_data:
        # Generate TTS audio for the slide
        if tts_enabled and slide_text:
            audio_file_name = image_path
            tts_audio_path = audio_file_name.replace(".png", ".mp3")
            asyncio.run(tts(slide_text, os.path.abspath(tts_audio_path), ai_voice=ai_voice))
            audio_clip = AudioFileClip(tts_audio_path)
            clip = ImageClip(image_path).set_duration(audio_clip.duration).resize(video_resolution)
            clip = clip.set_audio(audio_clip)
        else:
            clip = ImageClip(image_path).set_duration(5).resize(video_resolution)

        clips.append(clip)

    write_to_logs(f"-- Rendering Video...\n")
    # Concatenate all clips to create the video
    final_video = concatenate_videoclips(clips)
    final_video.write_videofile(
        output_video_path,
        fps=24,
        codec='libx264',
        audio_codec='aac',
        preset="slow"
    )
    write_to_logs(f"-- Video Render completed. Video path {output_video_path}\n")


def write_to_logs(message):
    log_text.config(state="normal")
    log_text.insert(tkinter.END, message)
    log_text.yview(tkinter.END)
    log_text.config(state="disabled")
    root.update_idletasks()


root = tkinter.Tk()
root.geometry("500x500")
root.title("PPT to Video")


label_width = 20
input_width = 24
label_font_options = tkFont.Font(family="Arial", size=12, weight=tkFont.NORMAL)
voice_names = ["en-US-AvaNeural", "en-US-BrianNeural", "en-IN-NeerjaNeural", "en-IN-PrabhatNeural", "en-US-AndrewNeural", "en-US-AriaNeural", "en-GB-LibbyNeural", "en-GB-RyanNeural"]
video_quality = ["480p", "720p", "1080p"]  # 720x480, 1280x720, 1920x1080

sv_ttk.set_theme("dark")

main_frame = ttk.Frame(root, padding="10")
main_frame.pack(fill="both", expand=True)

section1 = ttk.Labelframe(main_frame, text="Details", padding="10")
section1.pack(side="top", fill="both", expand=True, padx=5, pady=5)

ttk.Label(section1, text="Path to PPT:", width=label_width, font=label_font_options).grid(row=0, column=0, sticky="w", pady=5)
ppt_entry = ttk.Entry(section1, width=input_width)
ppt_entry.grid(row=0, column=1, pady=5)

select_ppt_button = ttk.Button(section1, text="...", command=select_input_ppt)
select_ppt_button.grid(row=0, column=2, padx=5, pady=5)

ttk.Label(section1, text="Select AI Voice:", width=label_width, font=label_font_options).grid(row=1, column=0, sticky="w", pady=5)
voice_combobox = ttk.Combobox(section1, values=voice_names, state="readonly")
voice_combobox.grid(row=1, column=1, pady=5)
voice_combobox.current(0)

ttk.Label(section1, text="Output Directory:", width=label_width, font=label_font_options).grid(row=2, column=0, sticky="w", pady=5)
output_dir_entry = ttk.Entry(section1, width=input_width)
output_dir_entry.grid(row=2, column=1, pady=5)

select_dir_button = ttk.Button(section1, text="...", command=select_output_directory)
select_dir_button.grid(row=2, column=2, padx=5, pady=5)

ttk.Label(section1, text="Select Video Quality:", width=label_width, font=label_font_options).grid(row=3, column=0, sticky="w", pady=5)
video_quality_combobox = ttk.Combobox(section1, values=video_quality, state="readonly")
video_quality_combobox.grid(row=3, column=1, pady=5)
video_quality_combobox.current(2)

ctv_button = ttk.Button(section1, text="Convert to Video", command=on_submit)
ctv_button.grid(row=4, column=0, columnspan=3, pady=10)

section2 = ttk.Labelframe(main_frame, text="Logs", padding="10")
section2.pack(side="top", fill="both", expand=True, padx=5, pady=5)

log_text = tkinter.Text(section2, wrap="word", height=10, state="disabled")
log_text.pack(fill="both", expand=True)

root.mainloop()

