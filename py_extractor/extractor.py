import os, zipfile, shutil
from pydub import AudioSegment
import speech_recognition as sr


audio_extensions = ('.aiff', '.au', '.mid', '.midi', '.mp3', '.m4a', '.mp4', '.wav', '.wma')
init_dir = os.getcwd()
init_files = os.listdir()
results_dir =  "../results"
ppt_dir = init_dir + '/ppt'
media_dir = ppt_dir + '/media'
output_file_prefix = "converted_"
result_dir_contents = os.listdir(results_dir)

def prepare_voice_file(path):
    audio_file = AudioSegment.from_file(path, format="m4a")
    wav_file = os.path.splitext(path)[0] + '.wav'
    audio_file.export(wav_file, format='wav')
    return wav_file

def main():
    print('Extracting...')

    if os.path.isdir(ppt_dir):
        shutil.rmtree(ppt_dir)
    
    for power in init_files :

        fbase = "".join(power.split(".")[:-1]).replace(' ', '_')

        if not power.endswith('.zip') and not power.endswith('.pptx') or power == 'base_library.zip':
            continue
        elif fbase in result_dir_contents:
            print(f"{fbase} already processed! Skipping...")
            continue

        base = os.path.splitext(power)[0]
        power_copy_copy = base + '_copy.pptx'
        new_zip = base + '.zip'
        shutil.copy(power, power_copy_copy)
        os.rename(power_copy_copy , new_zip)
        
        with zipfile.ZipFile(new_zip) as myzip:
            if myzip.testzip() != None:
                print("Some of your media files are corrupted")

            else :
                for file in myzip.namelist() :
                    if file.endswith(audio_extensions):
                        myzip.extract(file)
                
        myzip.close()
        os.remove(new_zip)
        break

    audio_files = os.listdir(media_dir)
    audio_file_paths = [f"{media_dir}/{f}" for f in audio_files]

    if audio_file_paths:
        r = sr.Recognizer()
        total_num = len(audio_file_paths)
        processing_count = 1

        result_dir_relative = "".join(power.split(".")[:-1]).replace(" ", "_")
        result_subdir = f"{results_dir}/{result_dir_relative}"
        try:
            os.mkdir(result_subdir)
        except: pass

        shutil.copy(power, result_subdir)

        for f in audio_file_paths:
            f_name = f.split("/")[-1].split(".")[0]
            print(f"Processing {processing_count}/{total_num}...")
            wav_file = prepare_voice_file(f)
            with sr.AudioFile(wav_file) as source:
                audio_data = r.record(source)
                text = r.recognize_google(audio_data)
                with open(f"{result_subdir}/{output_file_prefix}{f_name}.txt", 'w') as out_file:
                    out_file.write(text)
            processing_count += 1
    
    shutil.rmtree(media_dir)

if __name__ == "__main__":
    main()
