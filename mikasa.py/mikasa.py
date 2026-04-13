import time

import pyttsx3


def build_engine(retries=5, retry_delay=1):
    last_error = None
    for _ in range(retries):
        try:
            engine = pyttsx3.init()
            engine.setProperty("rate", 170)
            return engine
        except Exception as exc:
            last_error = exc
            time.sleep(retry_delay)

    raise RuntimeError(f"Could not initialize text-to-speech: {last_error}")


def main():
    engine = build_engine()
    engine.say("Hi Seenu, I am Mikasa. How can I help you?")
    engine.runAndWait()


if __name__ == "__main__":
    main()
