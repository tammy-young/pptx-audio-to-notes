# Powerpoint Audio Extractor
Extract audio from powerpoints to text

# Prerequisites
Requires Python 3. Then install the requirements.
```
pip install -r requirements.txt
```

# Usage
1. Copy a .pptx file into the `py_extractor` directory
2. `python extractor.py`
3. The extracted audio files can be found in `results/<PPTX_FILE_NAME>`
4. After `extractor.py` finishes, in the extracted audio files directory will be `joinall.py`
5. `python joinall.py`
6. `joined.txt` contains the transcribed audio, numbered with the corresponding slide number

# Inconsistent Audio
If the powerpoint doesn't have audio on every slide, you will have to add the slide numbers to the `skip` list in `joinall.py`. There is currently no automation for this.
