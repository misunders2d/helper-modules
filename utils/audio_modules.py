from pydub import AudioSegment
import os, pyperclip, pickle
from openai import OpenAI
from moviepy.editor import VideoFileClip
import re
import customtkinter as ctk
from concurrent.futures import ThreadPoolExecutor

from dotenv import load_dotenv

load_dotenv()

AI_KEY = os.getenv('AI_KEY_EMBEDDING')
client = OpenAI(api_key = AI_KEY)


out_path = os.path.join(os.path.expanduser('~'),'temp/pics/chunks')
parent_path = os.path.dirname(out_path)
temp_transcription_path = os.path.join(parent_path,'transcriptions_temp.pkl')
temp_transcription_path2 = os.path.join(parent_path,'transcriptions_temp.txt')
temp_summary_path = os.path.join(parent_path,'summaries_temp.txt')

if not os.path.exists(out_path):
    os.makedirs(out_path)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.geometry('400x300')
        self.title('Transcribe and summarize videos')
        
        self.executor = ThreadPoolExecutor()
        self.extract_button = ctk.CTkButton(self, text='1. Extract audio', command=self.run_extraction)
        self.extract_button.pack(pady=10)
        self.split_button = ctk.CTkButton(self, text='2. Split audio', command=self.split_audio_on_silence)
        self.split_button.pack(pady=10)
        self.transcribe_button = ctk.CTkButton(self, text='3. Transcribe audio', command=self.create_transcription)
        self.transcribe_button.pack(pady=10)
        self.summarize_button = ctk.CTkButton(self, text='4. Summarize', command=lambda: self.summarize_text(text))
        self.summarize_button.pack(pady=10)
        self.progress = ctk.CTkProgressBar(self, mode='indeterminate')
        self.progress.pack(pady=10)
        self.status_label = ctk.CTkLabel(self, text="")
        self.status_label.pack(pady=10)

    # STEP 1, EXTRACT AUDIO
    def run_extraction(self):
        self.executor.submit(self.extract_audio_from_video)
    
    def extract_audio_from_video(self):
        self.status_label.configure(text="audio extraction started")
        # video_file = sg.PopupGetFile('select video file')
        video_file = ctk.filedialog.askopenfile(title='Path to video file', initialdir=os.path.join(os.path.expanduser('~'),'Downloads'))
        if not video_file:
            return
        self.progress.start()
        video = VideoFileClip(video_file.name)
        
        audio = video.audio
        
        audio.write_audiofile(os.path.join(parent_path,'audio.mp3'))
        
        video.close()
        audio.close()
        self.status_label.configure(text="audio extraction completed")
        self.progress.stop()
    
    # STEP 2, SPLIT AUDIO TO CHUNKS OF ABOUT 15 MINUTES
    def split_audio_on_silence(self, file_path=os.path.join(parent_path,'audio.mp3'), max_size_mb=20, silence_thresh=-40):
        
        def estimate_mp3_duration(size_mb, bitrate_kbps=192):
            """Estimate duration in seconds that corresponds to a given size in MB at a specific bitrate."""
            return (size_mb * 1024 * 1024 * 8) / (bitrate_kbps * 1024)  # Convert MB to bits and divide by bitrate
    
        audio = AudioSegment.from_file(file_path, format="mp3")
        
        # Maximum target duration in seconds to reach the MP3 size limit
        # target_duration = estimate_mp3_duration(max_size_mb)
        target_duration = 900
        
        current_chunk = AudioSegment.silent(duration=0)
    
        for i, start in enumerate(range(0, len(audio), 1000)):  # Check every second
            end = start + 1000
            
            # Extract a 1-second chunk
            chunk = audio[start:end]
            
            # Check if the chunk is below the silence threshold
            if chunk.dBFS < silence_thresh:
                # If current chunk has audio, decide if it should be exported
                if len(current_chunk) > 0:
                    # If current chunk exceeds or reaches the target duration, save it
                    if len(current_chunk) / 1000 >= target_duration:
                        filename = os.path.join(out_path, f'chunk_{i}.mp3')
                        current_chunk.export(filename, format='mp3')
                        current_chunk = AudioSegment.silent(duration=0)  # Reset current chunk
            else:
                current_chunk += chunk  # Add to current non-silent chunk
            
            # If the current chunk length exceeds target duration, export it
            if len(current_chunk) / 1000 >= target_duration:
                filename = os.path.join(out_path, f'chunk_{i}.mp3')
                print(f'saving {filename}')
                current_chunk.export(filename, format='mp3')
                current_chunk = AudioSegment.silent(duration=0)  # Reset for a new chunk
    
        # Export any remaining audio
        if len(current_chunk) > 0:
            filename = os.path.join(out_path, f'chunk_{i}.mp3')
            print(f'saving {filename}')
            current_chunk.export(filename, format='mp3')
    
    # STEP 3, TRANSCRIBE AUDIO
    def create_transcription(self):
        global text, transcription_raw
        def transcribe_audio(audio_file_path):
            with open(audio_file_path, 'rb') as audio_file:
                transcription = client.audio.transcriptions.create(
                    model = "whisper-1",
                    file = audio_file,
                    response_format='verbose_json',
                    timestamp_granularities='segment')
            return transcription
        
        def convert_seconds(seconds):
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            remaining_seconds = seconds - (hours * 60 * 60) - (minutes * 60)
            return str(int(hours)).zfill(2), str(int(minutes)).zfill(2), str(int(remaining_seconds)).zfill(2)
        
        def process_transcription(transcription, offset: int):
            final_text = ''
            for segment in transcription.to_dict()['segments']:
                time = ':'.join(convert_seconds(segment['start'] + offset))
                text = segment['text']
                final_text += f'{time}: ' + text + '\n'
            offset += segment['end']
            return final_text, offset
            
        
        # folder = sg.PopupGetFolder('select folder with audio files', default_path=out_path, initial_folder=out_path)
        folder = ctk.filedialog.askdirectory(title='select folder with audio files', initialdir=out_path)
        if not folder:
            return
        
        files = sorted(
            [os.path.join(folder, x) for x in os.listdir(folder) if re.search('([0-9]+)',x)],
            key = lambda x: int(re.search('([0-9]+)',x).group()), reverse = False)
        
        transcriptions = {}
        transcriptions_raw = {}
        offset = 0
        try:
            for file in files:
                print(f'transcribing {file}')
                transcription = transcribe_audio(file)
                transcriptions_raw[file] = transcription
                t, offset = process_transcription(transcription, offset)
                transcriptions[file] = t
        except Exception as e:
            print(f'Exception occurred: {e}')
        finally:
            with open(temp_transcription_path, 'wb') as f:
                pickle.dump(transcriptions_raw, f)
        text = '\n\n'.join([value for value in transcriptions.values()])
        with open(temp_transcription_path2, 'w') as f:
            f.write(text)
        pyperclip.copy('RAW TRANSCRIPTION:\n'+text)
        print('Trascription copied to clipboard')
        # return text, transcriptions_raw
    
    # STEP 4, SUMMARIZE
    def summarize_text(self, text:str):
        pre_prompt = '''
        Please summarize the following transcription.
        Drop everything unrelated, but don't exclude any tools or hacks.
        Create a comprehensive list\
        of hacks/tools mentioned along with their use cases and benefits.\
        Include mentions and/or links to the tools/resources mentioned,\
        and also the problem that is being addressed with that hack or tool.\
        Also include any steps described, but don't expand too much, try\
        to stay within 2-3 paragraphs for each tool/hack mentioned.
        '''
        
        try:
            response = client.chat.completions.create(
            # model="gpt-4o-mini",
            model = "gpt-4o",
            messages=[
                {"role": "system", "content": "You help summarize and structure large texts and meeting transcriptions"},
                {"role": "user", "content": pre_prompt},
                {"role": "user", "content": text},
                ]
            )
        except Exception as e:
            print(f'Exception occurred: {e}')
        finally:
            with open(temp_summary_path, 'w') as f:
                f.write(response.choices[0].message.content)
        pyperclip.copy('Summary:\n'+response.choices[0].message.content)
        print('Summary copied to clipboard')
        return response.choices[0].message.content


if __name__ == "__main__":
    app = App()
    app.mainloop()