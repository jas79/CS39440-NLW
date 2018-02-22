import os;

path = '../ABERSHIP_transcription_vtls004566921'

for result in os.walk(path):
    if result[2].startswith("File"):
        print(result);
