import face_recognition
import cv2
import numpy as np
from win32com.client import Dispatch
import smtplib
import email
import imaplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import traceback
import time
import datetime
# from datetime import datetime
# import os


video_capture = cv2.VideoCapture(0)

def markAttendance(name):
    with open('Attendance.csv','r+') as f:
        mydatalist =f.readlines()
        namelist=[]
        for line in mydatalist:
            entry=line.split(',')
            namelist.append(entry[0])
        if name not in namelist:
            time_now = datetime.now()
            tStr=time_now.strftime('%H:%M:%S')
            dStr =time_now.strftime('%d:%m:%y')
            f.writelines(f'\n{name},{dStr},{tStr}')

def speak(str):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str)

def videorecord():
    face_cascade = cv2.CascadeClassifier(
        cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    body_cascade = cv2.CascadeClassifier(
        cv2.data.haarcascades + "haarcascade_fullbody.xml")
    detection = False
    detection_stopped_time = None
    timer_started = False
    SECONDS_TO_RECORD_AFTER_DETECTION = 5

    frame_size = (int(video_capture.get(3)), int(video_capture.get(4)))
    fourcc = cv2.VideoWriter_fourcc(*"mp4v")

    while True:
        _, frame = video_capture.read()

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, 1.3, 5)
        bodies = face_cascade.detectMultiScale(gray, 1.3, 5)

        if len(faces) + len(bodies) > 0:
            if detection:
                timer_started = False
            else:
                detection = True
                current_time = datetime.datetime.now().strftime("%d-%m-%Y-%H-%M-%S")
                out = cv2.VideoWriter(
                    f"{current_time}.mp4", fourcc, 20, frame_size)
                print("Started Recording!")
        elif detection:
            if timer_started:
                if time.time() - detection_stopped_time >= SECONDS_TO_RECORD_AFTER_DETECTION:
                    detection = False
                    timer_started = False
                    out.release()
                    print('Stop Recording!')
            else:
                timer_started = True
                detection_stopped_time = time.time()

        if detection:
            out.write(frame)

        # for (x, y, width, height) in faces:
        #    cv2.rectangle(frame, (x, y), (x + width, y + height), (255, 0, 0), 3)

        cv2.imshow("Camera", frame)

        if cv2.waitKey(1) == ord('q'):
            break

    out.release()

# read email details
ORG_EMAIL = "@dtu.ac.in"
FROM_EMAIL = "sachinmishra_ec20a14_65" + ORG_EMAIL
FROM_PWD = "Manya@3251"
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT = 993

def read_email_from_gmail():
    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)

        mail.select('inbox')
        data = mail.search(None, 'ALL')
        mail_ids = data[1]
        id_list = mail_ids[0].split()
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        for (index, i) in enumerate(range(latest_email_id,first_email_id, -1)):
            data = mail.fetch(str(i), '(RFC822)' )
            for response_part in data:
                arr = response_part[0]
                if isinstance(arr, tuple):
                    msg = email.message_from_string(str(arr[1],'utf-8'))
                    email_subject = msg['subject']
                    email_from = msg['from']
                    if email_subject == "A":
                        speak("Welcome guest you are now allowed to enter")
                        print("Allowed")
                    # print('From : ' + email_from + '\n')
                    # print('Subject : ' + email_subject + '\n')
                    elif email_subject == "D":
                        speak("You have been denied, this is the last warning to go back")
                        print("Denied")
                    else:
                        speak("you got no response from the authority, so you are not allowed")
            if index == 0:
                break
    except Exception as e:
        traceback.print_exc()
        print(str(e))

# video_capture = cv2.VideoCapture(0)

imgAttariak = face_recognition.load_image_file('1.jpg')
imgAttariak_encoding = face_recognition.face_encodings(imgAttariak)[0]

imgAli = face_recognition.load_image_file('2.jpg')
imgAli_encoding = face_recognition.face_encodings(imgAli)[0]

imgSachin = face_recognition.load_image_file('3.jpeg')
imgSachin_encoding = face_recognition.face_encodings(imgSachin)[0]

# Create arrays of known face encodings and their names
known_face_encodings = [
    imgAttariak_encoding,
    imgAli_encoding,
    imgSachin_encoding,
]
known_face_names = [
    "Jeetendra",
    "Deepak",
    "Sachin"
]

# Initialize some variables
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True
i = 1
while True and i==1:
    # Grab a single frame of video
    ret, frame = video_capture.read()

    # Resize frame of video to 1/4 size for faster face recognition processing
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

    # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
    rgb_small_frame = small_frame[:, :, ::-1]

    # Only process every other frame of video to save time
    if process_this_frame:
        # Find all the faces and face encodings in the current frame of video
        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

        face_names = []
        for face_encoding in face_encodings:
            # See if the face is a match for the known face(s)
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            name = "Unknown"

            # # If a match was found in known_face_encodings, just use the first one.
            # if True in matches:
            # Or instead, use the known face with the smallest distance to the new face
            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)
            if matches[best_match_index]:
                name = known_face_names[best_match_index]
                speak("Welcome" + name)
                markAttendance(name)
                videorecord()
                i += 1
            else:
                speak("Go from here now")
                # cam = cv2.VideoCapture(0)
                # cv2.namedWindow("test")
                img_counter = 1
                while True and img_counter==1:
                    ret, frame = video_capture.read()
                    if not ret:
                        print("failed to grab frame")
                        break
                    cv2.imshow("test", frame)
                    img_name = "Unknown_{}.png".format(img_counter)
                    cv2.imwrite(img_name, frame)
                    print("{} written!".format(img_name))

                    fromaddr = "sachinmishra_ec20a14_65@dtu.ac.in"
                    password = "Manya@3251"
                    toaddr = "bhautaalsi@gmail.com"
                    # instance of MIMEMultipart
                    msg = MIMEMultipart()
                    # storing the senders email address
                    msg['From'] = fromaddr
                    msg['To'] = toaddr
                    msg['Subject'] = "Someone is at door"

                    body = "Send A for allowed and D for deny"
                    msg.attach(MIMEText(body, 'plain'))

                    filename = "Unknown_1.png"
                    attachment = open(filename, "rb")

                    # instance of MIMEBase and named as p
                    p = MIMEBase('application', 'octet-stream')

                    # To change the payload into encoded form
                    p.set_payload((attachment).read())

                    # encode into base64
                    encoders.encode_base64(p)

                    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

                    # attach the instance 'p' to instance 'msg'
                    msg.attach(p)

                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(fromaddr, password)
                    msg.as_string()
                    server.send_message(msg)
                    server.quit()
                    img_counter += 1
                    k = cv2.waitKey(1)

                # cam.release()
                # cv2.destroyAllWindows()
                time.sleep(40)
                read_email_from_gmail()
                # speak("Go from here now, this is the last chance")
            face_names.append(name)

    process_this_frame = not process_this_frame


    # Display the results
    for (top, right, bottom, left), name in zip(face_locations, face_names):
        # Scale back up face locations since the frame we detected in was scaled to 1/4 size
        top *= 4
        right *= 4
        bottom *= 4
        left *= 4

        # Draw a box around the face
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)

        # Draw a label with a name below the face
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

    # Display the resulting image
    cv2.imshow('Video', frame)

    # Hit 'q' on the keyboard to quit!
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Release handle to the webcam
video_capture.release()
cv2.destroyAllWindows()