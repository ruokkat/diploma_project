import cv2
import mediapipe as mp
import keyboard
import pyautogui
import time
from collections import deque
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume_controller = cast(interface, POINTER(IAudioEndpointVolume))
last_volume_level = volume_controller.GetMasterVolumeLevelScalar()


mp_hands = mp.solutions.hands
mp_drawing = mp.solutions.drawing_utils

cap = cv2.VideoCapture(0)
cap.set(3, 640)
cap.set(4, 480)

gesture_buffer = deque(maxlen=5)
prev_action_time = 0
min_delay = 2

def count_fingers(hand_landmarks):
    tips_ids = [4, 8, 12, 16, 20]
    fingers = []
    if hand_landmarks.landmark[tips_ids[0]].x < hand_landmarks.landmark[tips_ids[0] - 1].x:
        fingers.append(1)
    else:
        fingers.append(0)
    for i in range(1, 5):
        if hand_landmarks.landmark[tips_ids[i]].y < hand_landmarks.landmark[tips_ids[i] - 2].y:
            fingers.append(1)
        else:
            fingers.append(0)
    return fingers

def detect_gesture(fingers):
    if fingers == [0, 1, 0, 0, 0]:
        return 'one_finger'
    elif fingers == [0, 1, 1, 0, 0]:
        return 'two_fingers'
    elif fingers == [1, 1, 1, 1, 1]:
        return 'open_palm'
    elif fingers == [0, 0, 0, 0, 0]:
        return 'fist'
    return 'unknown'

def most_common(buffer):
    return max(set(buffer), key=buffer.count)

def adjust_volume_by_hand_position(hand_landmarks):
    global last_volume_level
    y_coords = [lm.y for lm in hand_landmarks.landmark]
    avg_y = sum(y_coords) / len(y_coords)
    volume_level = max(0.0, min(1.0, 1.3 - avg_y * 2))
    if abs(volume_level - last_volume_level) > 0.02:
        volume_controller.SetMasterVolumeLevelScalar(volume_level, None)
        print(f"🔊 Volume: {int(volume_level * 100)}%")
        last_volume_level = volume_level


with mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.5) as hands:
    while cap.isOpened():
        success, frame = cap.read()
        if not success:
            break

        frame = cv2.flip(frame, 1)
        image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = hands.process(image)
        image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

        gesture = 'none'

        if results.multi_hand_landmarks:
            for hand_landmarks in results.multi_hand_landmarks:
                mp_drawing.draw_landmarks(image, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                fingers = count_fingers(hand_landmarks)
                gesture = detect_gesture(fingers)
                gesture_buffer.append(gesture)

                if gesture == 'open_palm':
                    adjust_volume_by_hand_position(hand_landmarks)

        if len(gesture_buffer) == gesture_buffer.maxlen:
            common_gesture = most_common(gesture_buffer)
            if common_gesture != 'unknown' and time.time() - prev_action_time > min_delay:
                if common_gesture == 'fist':
                    print("▶️⏸ Play/Pause")
                    keyboard.send('play/pause media')
                    pyautogui.press('k')
                elif common_gesture == 'one_finger':
                    print("⏭ Наступний трек")
                    keyboard.send('next track')
                    pyautogui.press('l')
                elif common_gesture == 'two_fingers':
                    print("⏮ Попередній трек")
                    keyboard.send('previous track')
                    pyautogui.press('j')
                prev_action_time = time.time()
                gesture_buffer.clear()

        cv2.imshow("Gesture Control", image)
        if cv2.waitKey(5) & 0xFF == 27:
            break

cap.release()
cv2.destroyAllWindows()
