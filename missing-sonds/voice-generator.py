import requests
import openpyxl

# Constants
CHUNK_SIZE = 1024  # Size of chunks to read/write at a time
XI_API_KEY = ""  # Your API key for authentication
VOICE_ID = ""  # ID of the voice model to use
XLSX_PATH = "audios.xlsx"  # Path to the input Excel file
MODEL_ID = "eleven_multilingual_v2"  # Model ID for the TTS
STABILITY = 0.5  # Stability setting for the voice
SIMILARITY_BOOST = 0.9  # Similarity boost setting for the voice
STYLE = 0.7  #velocity
USE_SPEAKER_BOOST = True  # Use speaker boost setting for the voice

# Construct the URL for the Text-to-Speech API request
tts_url = f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE_ID}/stream"

# Set up headers for the API request, including the API key for authentication
headers = {
    "Accept": "application/json",
    "xi-api-key": XI_API_KEY
}

# Read the Excel file
workbook = openpyxl.load_workbook(XLSX_PATH)
sheet = workbook.active

# Iterate over the rows in the Excel sheet
for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True):
    TEXT_TO_SPEAK = row[0]
    OUTPUT_PATH = f"{TEXT_TO_SPEAK}.mp3"

    # Set up the data payload for the API request
    data = {
        "text": TEXT_TO_SPEAK,
        "model_id": MODEL_ID,
        "voice_settings": {
            "stability": STABILITY,
            "similarity_boost": SIMILARITY_BOOST,
            "style": STYLE,
            "use_speaker_boost": USE_SPEAKER_BOOST
        }
    }

    # Make the POST request to the TTS API with headers and data, enabling streaming response
    response = requests.post(tts_url, headers=headers, json=data, stream=True)

    # Check if the request was successful
    if response.ok:
        # Open the output file in write-binary mode
        with open(OUTPUT_PATH, "wb") as f:
            # Read the response in chunks and write to the file
            for chunk in response.iter_content(chunk_size=CHUNK_SIZE):
                f.write(chunk)
        # Inform the user of success
        print(f"Audio stream for '{TEXT_TO_SPEAK}' saved successfully.")
    else:
        # Print the error message if the request was not successful
        print(f"Failed to generate audio for '{TEXT_TO_SPEAK}': {response.text}")
