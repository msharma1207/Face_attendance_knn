import sys
import cv2
import pickle
import os
import csv
import time
import numpy as np
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtCore import QTimer
from sklearn.neighbors import KNeighborsClassifier
from win32com.client import Dispatch

# Function for Text-to-Speech
def speak(text):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

class FaceAttendanceApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Face Attendance System")
        self.setGeometry(100, 100, 800, 600)

        # Load trained face data
        with open("data/names.pkl", "rb") as f:
            self.LABELS = pickle.load(f)
        with open("data/faces_data.pkl", "rb") as f:
            self.FACES = pickle.load(f)

        print("✅ Shape of Faces matrix -->", self.FACES.shape)
        
        self.knn = KNeighborsClassifier(n_neighbors=5)
        self.knn.fit(self.FACES, self.LABELS)
        
        self.video = cv2.VideoCapture(0)
        self.facedetect = cv2.CascadeClassifier("data/haarcascade_frontalface_default.xml")

        # UI Elements
        self.image_label = QLabel(self)
        self.start_button = QPushButton("Start Recognition")
        self.stop_button = QPushButton("Stop Recognition")
        self.attendance_table = QTableWidget()
        self.attendance_table.setColumnCount(2)
        self.attendance_table.setHorizontalHeaderLabels(["Name", "Time"])

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.image_label)
        layout.addWidget(self.start_button)
        layout.addWidget(self.stop_button)
        layout.addWidget(self.attendance_table)
        self.setLayout(layout)

        # Timer for Video Feed
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        
        # Button Actions
        self.start_button.clicked.connect(self.start_recognition)
        self.stop_button.clicked.connect(self.stop_recognition)
        
    def start_recognition(self):
        self.timer.start(30)

    def stop_recognition(self):
        self.timer.stop()
        self.video.release()

    def update_frame(self):
        ret, frame = self.video.read()
        if not ret:
            return

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = self.facedetect.detectMultiScale(gray, 1.3, 5)

        for (x, y, w, h) in faces:
            crop_img = frame[y:y+h, x:x+w, :]
            resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
            output = self.knn.predict(resized_img)

            ts = time.time()
            date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
            timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
            csv_filename = f"Attendance/Attendance_{date}.csv"
            exist = os.path.isfile(csv_filename)
            
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
            cv2.putText(frame, str(output[0]), (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 255, 255), 2)

            self.mark_attendance(output[0], timestamp, csv_filename, exist)

        self.display_frame(frame)

    def display_frame(self, frame):
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = frame.shape
        bytes_per_line = ch * w
        convert_to_qt_format = QImage(frame.data, w, h, bytes_per_line, QImage.Format_RGB888)
        self.image_label.setPixmap(QPixmap.fromImage(convert_to_qt_format))

    def mark_attendance(self, name, timestamp, csv_filename, exist):
        with open(csv_filename, "a", newline="") as csvfile:
            writer = csv.writer(csvfile)
            if not exist:
                writer.writerow(["NAME", "TIME"])
            writer.writerow([name, timestamp])
        self.update_table(name, timestamp)

    def update_table(self, name, timestamp):
        row_position = self.attendance_table.rowCount()
        self.attendance_table.insertRow(row_position)
        self.attendance_table.setItem(row_position, 0, QTableWidgetItem(name))
        self.attendance_table.setItem(row_position, 1, QTableWidgetItem(timestamp))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FaceAttendanceApp()
    window.show()
    sys.exit(app.exec_())