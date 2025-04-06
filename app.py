import streamlit as st
from pptx import Presentation
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip, concatenate_audioclips
import os
import shutil
import comtypes.client
import comtypes
from PIL import Image

st.title("🎥 PPT to Video Creator (Real Slides + Background Music)")

# Upload PPT
ppt_file = st.file_uploader("Upload your PPT file", type=["pptx"])

# Upload Audio
audio_file = st.file_uploader("Upload your Audio file", type=["mp3", "wav"])

# Enter Duration
duration_per_slide = st.number_input("Enter duration per slide (in seconds)", min_value=1, value=10)

# Start Button
if st.button("Create Video"):
    if ppt_file and audio_file:
        with st.spinner('Processing your files... Please wait...'):
            # Create temp folder safely
            temp_folder = "temp_slides"
            if os.path.exists(temp_folder):
                shutil.rmtree(temp_folder)
            os.makedirs(temp_folder)

            # Save uploaded PPT file
            ppt_path = "uploaded_ppt.pptx"
            with open(ppt_path, "wb") as f:
                f.write(ppt_file.read())

            # Initialize COM for PowerPoint Automation
            comtypes.CoInitialize()
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1

            # Open PPT and export slides as images
            deck = powerpoint.Presentations.Open(os.path.abspath(ppt_path))
            export_folder = os.path.abspath(temp_folder)
            deck.SaveAs(export_folder, 17)  # 17 = ppSaveAsJPG
            deck.Close()
            powerpoint.Quit()

            # Check if images exported
            image_files = sorted([
                os.path.join(temp_folder, img)
                for img in os.listdir(temp_folder)
                if img.lower().endswith(".jpg")
            ])

            if len(image_files) == 0:
                st.error("❗ No slides found after exporting PPT! Please check your PPT file.")
                raise SystemExit

            # Create video clips from exported images
            image_clips = []
            for img_path in image_files:
                clip = ImageClip(img_path).set_duration(duration_per_slide)
                image_clips.append(clip)

            final_clip = concatenate_videoclips(image_clips, method="compose")

            # Save uploaded audio
            audio_path = "background_audio.mp3"
            with open(audio_path, "wb") as f:
                f.write(audio_file.read())

            # Handle audio looping
            audio = AudioFileClip(audio_path)
            if audio.duration < final_clip.duration:
                n_loops = int(final_clip.duration // audio.duration) + 1
                audio = concatenate_audioclips([audio] * n_loops)
            audio = audio.subclip(0, final_clip.duration)

            final_clip = final_clip.set_audio(audio)

            # Export final video
            output_video = "final_video.mp4"
            final_clip.write_videofile(output_video, fps=24)

            # Offer download button
            with open(output_video, "rb") as video_file:
                st.download_button(
                    label="📥 Download Your Final Video",
                    data=video_file,
                    file_name="ppt_to_video.mp4",
                    mime="video/mp4"
                )

            # Clean up everything properly
            shutil.rmtree(temp_folder)
            os.remove(ppt_path)
            audio.close()  # 🔥 Close the audio file first!
            os.remove(audio_path)
            os.remove(output_video)

            st.success("✅ Video created successfully!")

    else:
        st.error("❗ Please upload both a PPT file and an Audio file.")
