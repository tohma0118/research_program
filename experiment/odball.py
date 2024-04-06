"""20230609作成 keisoku.pyの改良版

変更点
・高頻度刺激を赤→緑,低頻度刺激を青→赤に
・刺激提示の秒数を0.5秒→0.3秒に
・ターゲット（低頻度刺激）を数えさせるのではなく、表示中にボタンを押させる
logファイルはtargetlog.log
"""

# 刺激提示間隔3s
import random
import time
import serial
import tkinter as tk
import logging

LED = False  # LEDで外乱を与えたい場合はTrue,そうじゃないならFalse

BLINKS = 20  # LED提示回数 20230524 50→20
STIMULUS_PRESENTATIONS = 200  # 刺激提示回数

count = 0
ns = []  # ターゲットの番号
enter_count = 0
circle_state = ""  # 画面上の円の状態 hidden,green,redの3種類

logger = logging.getLogger("target")
logger.setLevel(10)
fh = logging.FileHandler("C:/python_program/research_program/experiment/targetlog.log")
logger.addHandler(fh)

formatter = logging.Formatter("%(asctime)s:%(levelname)s:%(message)s")
fh.setFormatter(formatter)


def closePort():
    global ser, root
    if LED:
        ser.close()
    root.destroy()  # 画面を消す


def lightOnOff():
    global ser
    for i in range(BLINKS):
        ser.write(b"1")  # Arduino に 1 を送る
        time.sleep(0.21)
        ser.write(b"0")  # LED を消す
        time.sleep(0.2)


def hidden():
    global canvas, root, circle_state

    canvas.itemconfig("oval", state=tk.HIDDEN)
    circle_state = "hidden"
    root.after(2700, normal)


def normal():
    global count, canvas, root, ns, circle_state, enter_count

    count += 1

    if count in ns:
        canvas.itemconfig("oval", state=tk.NORMAL, fill="red")
        circle_state = "red"
    else:
        canvas.itemconfig("oval", state=tk.NORMAL, fill="green")
        circle_state = "green"

    if count <= STIMULUS_PRESENTATIONS:
        root.after(300, hidden)
    else:
        count = 0
        label2.config(text=len(ns))  # 20230601 str(ns)→len(ns)
        label4.config(text=enter_count)

        logger.log(10, str(ns) + "\n")

        enter_count = 0
        ns.clear()
        canvas.itemconfig("oval", state=tk.HIDDEN)


def rand_ints_nodup(k):
    global ns
    while len(ns) < k:
        n = random.randint(4, STIMULUS_PRESENTATIONS)
        if all(
            num not in ns for num in (n - 1, n, n + 1)
        ):  # 20230601　値が連続しないように変更
            ns.append(n)
    ns.sort()


def prepare():
    global canvas, root, ns, label2

    k = random.randint(
        int(STIMULUS_PRESENTATIONS * 0.15), int(STIMULUS_PRESENTATIONS * 0.2)
    )

    rand_ints_nodup(k)

    # canvas.create_oval(180, 180, 420, 420, fill="red", tag="oval")
    canvas.create_oval(50, 50, 750, 750, fill="green", tag="oval")  # 20230601 変更

    if LED:
        lightOnOff()

    hidden()

    label2.config(text="")
    label4.config(text="")


def count_enter_key(event):
    global enter_count, circle_state
    if event.keysym == "Return":
        enter_count += 1
    logger.log(20, "enter pressed. circle:%s", circle_state)


# 画面構築
root = tk.Tk()
root.title("circle")
root.geometry("100x100")  # 画面サイズは 1920x1080

if LED:
    ser = serial.Serial("COM3", 9600)

canvas = tk.Canvas(
    root,
    width=800,
    height=800,
)  # 20230601 widthとheightを600から800に変更
canvas.pack()

button = tk.Button(root, text="start", height=10, width=20, command=prepare)
button.place(relx=0.9, rely=0.8, anchor=tk.SE)

label = tk.Label(root, text="target:", font=(" ", 50))
label.place(relx=0.1, rely=0.75, anchor=tk.SW)

label2 = tk.Label(root, text="", font=(" ", 50))
label2.place(relx=0.23, rely=0.75, anchor=tk.SW)

label3 = tk.Label(root, text="enter:", font=(" ", 50))
label3.place(relx=0.1, rely=0.85, anchor=tk.SW)

label4 = tk.Label(root, text="", font=(" ", 50))
label4.place(relx=0.22, rely=0.85, anchor=tk.SW)

root.bind("<Key>", count_enter_key)
root.focus_set()

root.protocol("WM_DELETE_WINDOW", closePort)  # 罰ボタンが押されたら closePort を実行
root.mainloop()
