import cv2
import pickle
import os
import csv
import time
from datetime import datetime
from sklearn.neighbors import KNeighborsClassifier
from win32com.client import Dispatch

# 🎤 Function for Text-to-Speech
def speak(text):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

# 📷 Initialize Webcam & Face Detection
video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier("data/haarcascade_frontalface_default.xml")

# 🔍 Load Face Data & Labels
try:
    with open("data/names.pkl", "rb") as f:
        LABELS = pickle.load(f)

    with open("data/faces_data.pkl", "rb") as f:
        FACES = pickle.load(f)

    print("✅ Shape of Faces matrix -->", FACES.shape)

    # 🧠 Train KNN Model
    knn = KNeighborsClassifier(n_neighbors=5)
    knn.fit(FACES, LABELS)

except FileNotFoundError:
    print("❌ Error: Missing face data files (names.pkl or faces_data.pkl). Train the model first.")
    exit()

# 🎨 Load Background Image
background_path = r"C:\Users\User\Desktop\face_attendence\background_frame.png"

# Ensure the file exists
if not os.path.isfile(background_path):
    print(f"❌ Error: Background image not found! Check the path: {background_path}")
    print(f"📂 Files in directory: {os.listdir(os.path.dirname(background_path))}")
    exit()

# Load the image
imgBackground = cv2.imread(background_path)

# Ensure OpenCV successfully loads the image
if imgBackground is None:
    print("❌ Error: OpenCV could not read the image. Ensure it's a valid PNG file.")
    exit()

print("✅ Background image loaded successfully!")

# 📂 Ensure 'Attendance' Folder Exists
attendance_dir = "Attendance"
if not os.path.exists(attendance_dir):
    os.makedirs(attendance_dir)  # Create the folder if it doesn't exist
    print(f"📁 Created folder: {attendance_dir}")

# 🗂 CSV Columns for Attendance
COL_NAMES = ["NAME", "TIME"]

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)

        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        csv_filename = os.path.join(attendance_dir, f"Attendance_{date}.csv")
        exist = os.path.isfile(csv_filename)

        # 🏷 Draw Rectangle and Label
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)

        attendance = [str(output[0]), str(timestamp)]

    # 🖼 Overlay Frame on Background Image
    imgBackground[162:162+480, 55:55+640] = frame
    cv2.imshow("Attendance System", imgBackground)

    k = cv2.waitKey(1)
    
    # 📌 Mark Attendance on 'O' Press
    if k == ord("o"):
        speak("Ram Ram G .. Attendance Taken..")
        time.sleep(2)
        
        with open(csv_filename, "a", newline="") as csvfile:
            writer = csv.writer(csvfile)
            if not exist:
                writer.writerow(COL_NAMES)  # Write headers if file is new
            writer.writerow(attendance)

    # 🔴 Quit on 'Q' Press
    if k == ord("q"):
        break

video.release()
cv2.destroyAllWindows()
