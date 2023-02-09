import os
import cv2
import time
import shutil
import brickpi3

cam = cv2.VideoCapture(0)
cam.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc("H", "2", "6", "4"))
cam.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
cam.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

BP = brickpi3.BrickPi3()
BP.set_sensor_type(BP.PORT_1, BP.SENSOR_TYPE.TOUCH)
BP.set_sensor_type(BP.PORT_2, BP.SENSOR_TYPE.TOUCH)


BP.set_led(0)

target_dir = "/home/pi/testphoto"
shutil.rmtree(target_dir)
os.mkdir(target_dir)
print(cam)
n = 0

while True:
    try:
        isStart = BP.get_sensor(BP.PORT_1)
        if isStart == 1:
            print("start")
            n += 1
            counter = 0
            os.mkdir(f"{target_dir}/{n}")
            start_time = time.time()
            BP.set_led(99)
            with open(f"{target_dir}/{n}/{n}.txt", "w") as f:
                while True:
                    shatter = BP.get_sensor(BP.PORT_2)
                    if shatter == 1:
                        print("shatter")
                        counter += 1
                        ret, frame = cam.read()
                        cv2.imwrite(f"{target_dir}/{n}/{counter}.jpg", frame)
                        time.sleep(0.1)
                        BP.set_led(10)
                        time.sleep(0.1)
                        BP.set_led(99)
                        f.write(f"{time.time()-start_time},")
                    time.sleep(0.3)
                    isStart = BP.get_sensor(BP.PORT_1)
                    if isStart == 1:
                        print("stop")
                        f.write(f"{time.time()-start_time}")
                        BP.set_led(0)
                        time.sleep(0.3)
                        break
                # else:
                #     print("ss")
                #     continue
                # break
    except KeyboardInterrupt:
        print("Ctrl+Cでcameraは停止しました")
        # save
        BP.reset_all()
        break

# cam.release()
