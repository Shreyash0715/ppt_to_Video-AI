import streamlit as st
import os
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip, concatenate_audioclips
from pdf2image import convert_from_path
from PIL import Image
import tempfile

st.title("🎥 PPT to Video Creator (Real Slides + Background Music)")

# Upload PPTX
ppt_file = st.file_uploader("Upload your PPT file (.pptx)", type=["pptx"])

# Upload Audio
audio_file = st.file_uploader("Upload your Background Audio file (.mp3 or .wav)", type=["mp3", "wav"])

# Enter Duration
duration_per_slide = st.number_input("Enter Duration per Slide (in seconds)", min_value=1, value=5)

# Create Video Button
if st.button("Create Video"):
    if ppt_file and audio_file:
        with st.spinner('Processing your files... ⏳'):

            # Save the uploaded PPTX
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt:
                tmp_ppt.write(ppt_file.read())
                ppt_path = tmp_ppt.name

            # Convert PPTX to PDF using libreoffice (pre-installed on Streamlit Cloud)
            pdf_path = ppt_path.replace(".pptx", ".pdf")
            os.system(f"libreoffice --headless --convert-to pdf {ppt_path} --outdir {os.path.dirname(ppt_path)}")

            # Convert PDF pages to images
            slides = convert_from_path(pdf_path)
            temp_dir = tempfile.mkdtemp()
            image_paths = []

            for idx, slide in enumerate(slides):
                img_path = os.path.join(temp_dir, f"slide_{idx}.png")
                slide.save(img_path, "PNG")
                image_paths.append(img_path)

            # Create video from images
            clips = []
            for img_path in image_paths:
                img_clip = ImageClip(img_path).set_duration(duration_per_slide)
                clips.append(img_clip)

            final_video = concatenate_videoclips(clips, method="compose")

            # Save uploaded audio
            with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmp_audio:
                tmp_audio.write(audio_file.read())
                audio_path = tmp_audio.name

            # Add background audio (loop if needed)
            audio_clip = AudioFileClip(audio_path)
            if audio_clip.duration < final_video.duration:
                loops = int(final_video.duration // audio_clip.duration) + 1
                audio_clip = concatenate_audioclips([audio_clip] * loops)

            audio_clip = audio_clip.subclip(0, final_video.duration)
            final_video = final_video.set_audio(audio_clip)

            # Export Final Video
            output_path = os.path.join(temp_dir, "final_video.mp4")
            final_video.write_videofile(output_path, fps=24, codec='libx264')

            # Download button
            with open(output_path, "rb") as file:
                st.download_button(
                    label="📥 Download Your Final Video",
                    data=file,
                    file_name="ppt_to_video.mp4",
                    mime="video/mp4"
                )

            # Clean up temporary files
            try:
                os.remove(ppt_path)
                os.remove(pdf_path)
                os.remove(audio_path)
                for img in image_paths:
                    os.remove(img)
                os.remove(output_path)
            except:
                pass

            st.success("✅ Your video has been created successfully!")

    else:
        st.error("❗ Please upload both a PPT file and an Audio file before creating the video.")
