
from flask import Flask, render_template, flash, request, session
from flask import render_template, redirect, url_for, request



import smtplib



app = Flask(__name__)
app.config.from_object(__name__)
app.config['SECRET_KEY'] = '7d441f27d441f27567d441f2b6176a'

app.config['DEBUG']


@app.route("/")
def homepage():
    return render_template('index.html')
@app.route("/UserLogin")
def UserLogin():
    return render_template('UserLogin.html')


@app.route("/start", methods=['GET', 'POST'])
def start():
    error = None
    if request.method == 'POST':
        import csv
        import copy
        import cv2 as cv
        import mediapipe as mp
        from model import KeyPointClassifier
        from app_files import calc_landmark_list, draw_info_text, draw_landmarks, get_args, pre_process_landmark
        from PIL import Image, ImageDraw, ImageFont
        import numpy as np


        args = get_args()
        cap_device = args.device
        cap_width = args.width
        cap_height = args.height

        use_static_image_mode = args.use_static_image_mode
        min_detection_confidence = args.min_detection_confidence
        min_tracking_confidence = args.min_tracking_confidence

        cap = cv.VideoCapture(cap_device)
        cap.set(cv.CAP_PROP_FRAME_WIDTH, cap_width)
        cap.set(cv.CAP_PROP_FRAME_HEIGHT, cap_height)

        mp_hands = mp.solutions.hands
        hands = mp_hands.Hands(
            static_image_mode=use_static_image_mode,
            max_num_hands=1,
            min_detection_confidence=min_detection_confidence,
            min_tracking_confidence=min_tracking_confidence,
        )

        keypoint_classifier = KeyPointClassifier()

        with open('model/keypoint_classifier/keypoint_classifier_label.csv', encoding='utf-8-sig') as f:
            keypoint_classifier_labels = csv.reader(f)
            keypoint_classifier_labels = [
                row[0] for row in keypoint_classifier_labels
            ]

        flag = 0

        import win32com.client as wincl
        speak = wincl.Dispatch("SAPI.SpVoice")

        while True:
            key = cv.waitKey(10)
            if key == 27:  # ESC
                break

            ret, image = cap.read()
            if not ret:
                break
            image = cv.flip(image, 1)
            debug_image = copy.deepcopy(image)
            # print(debug_image.shape)
            # cv.imshow("debug_image",debug_image)
            image = cv.cvtColor(image, cv.COLOR_BGR2RGB)

            image.flags.writeable = False
            results = hands.process(image)
            image.flags.writeable = True

            if results.multi_hand_landmarks is not None:
                for hand_landmarks, handedness in zip(results.multi_hand_landmarks, results.multi_handedness):
                    landmark_list = calc_landmark_list(debug_image, hand_landmarks)

                    # print(hand_landmarks)
                    pre_processed_landmark_list = pre_process_landmark(landmark_list)

                    hand_sign_id = keypoint_classifier(pre_processed_landmark_list)

                    debug_image = draw_landmarks(debug_image, landmark_list)
                    flag += 1
                    print(flag)
                    if (flag == 100):
                        flag = 0

                        speak.Speak(keypoint_classifier_labels[hand_sign_id])

                    debug_image = draw_info_text(
                        debug_image,
                        handedness,
                        keypoint_classifier_labels[hand_sign_id])

            cv.imshow('Hand Gesture Recognition', debug_image)

        cap.release()
        cv.destroyAllWindows()




    return render_template('UserLogin.html')













if __name__ == '__main__':
    app.run(debug=True, use_reloader=True)
