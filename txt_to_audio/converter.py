import pyttsx3


# If you receive errors such as No module named win32com.client,
# No module named win32, or No module named win32api, you will need to additionally install pypiwin32.


class TextToAudio:
    engine: pyttsx3.Engine

    def __init__(self, voice, rate: int, volume: float):
        self.engine = pyttsx3.init()
        if voice:
            self.engine.setProperty('voice', voice)
        self.engine.setProperty('rate', rate)  # default 200
        self.engine.setProperty('volume', volume)  # between 0 and 1.0

    def list_available_voices(self):
        voices: list = [self.engine.getProperty('voices')]

        for i, voice in enumerate(voices[0]):
            print(f'{i + 1} {voice.name} {voice.age}: {voice.languages} ({voice.gender}) [ID: {voice.id}]')

    def text_to_audio(self, text: str, save: bool = False, filename='output.mp3'):
        self.engine.say(text)
        print('I am speaking')

        if save:
            self.engine.save_to_file(text, filename)

        self.engine.runAndWait()


# HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_RU-RU_IRINA_11.0 - Для русского языка
# HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0 - For english language


if __name__ == '__main__':
    tts = TextToAudio(None, 200, 1.0)
    #    tts.list_available_voices()
    tts.text_to_audio('Hello user', save=False, filename='output.mp3')
