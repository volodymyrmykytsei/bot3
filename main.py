import ImageGrab
import Image
import os
import time
import win32api, win32con

jew = [(76, 125, 169), (82, 76, 76), (61, 60, 64), (78, 124, 161), (79, 111, 131),
       (80, 83, 93), (73, 122, 176), (116, 120, 115), (80, 83, 93)]
take_pixel = (710,190)
take_color = (123, 68, 19)
take_action = (680,610)
close_jew_action = (990,80)
zoom_in_action = (1165,180)
energy_pointer = (975,45)
energy_pointer_color_if_full = (71, 206, 220)
open_store_action = (1060,740)
energy_item_action = (400, 445)
energy_item_color = (0,255,255)
close_store_action = (975,95)
open_jew_action = (668,173)
arrow_action = (1032,415)
start_game_action = (690,490)

y_edge = 147
x_edge = 445
blue_color = (2,145,156)
green_color = (89,165,79)
yellow_color = (250,211,54)
red_color = (156,96,89)
purple_color = (155,74,196)
cyan_color = (69, 108, 217)
BLUE = 1
GREEN = 2
YELLOW = 3
RED = 4
PURPLE = 5
CYAN = 6
matr = [[0 for i in range(7)] for j in range(7)]
matr1 = [[0 for i in range(7)] for j in range(7)]
b = [[0 for i in range(7)] for j in range(7)]
b1 = [[0 for i in range(7)] for j in range(7)]
def getCoord(x, y):
    return (x*80+x_edge, y*80+y_edge)

def mousePos(cord):
    win32api.SetCursorPos((cord[0], cord[1]))

def leftDown():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(.1)

def leftUp():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
    time.sleep(.1)

def mouseDrag(c1, c2):
    mousePos(c1)
    leftDown()
    leftUp()
    mousePos(c2)
    leftDown()
    leftUp()

def makeMove(x1, y1, x2, y2):
    mouseDrag(getCoord(x1,y1), getCoord(x2,y2))

def makeScreens():
    for i in range(100):
        time.sleep(1)
        im = ImageGrab.grab()
        im.save(os.getcwd() + '\\full_snap__' + str(int(time.time())) +
    '.png', 'PNG')

def screenGrab():
    box = ()
    im = ImageGrab.grab()
    return im;

def getColor1(x):
    if (x == blue_color):
        return BLUE
    if (x == red_color):
        return RED
    if (x == yellow_color):
        return YELLOW
    if (x == green_color):
        return GREEN
    if (x == purple_color):
        return PURPLE
    if (x == cyan_color):
        return CYAN
    return 0

def getColor(x, im, x1, y1):
    if (x == blue_color):
        return BLUE
    if (x == red_color):
        return RED
    if (x == yellow_color):
        return YELLOW
    if (x == green_color):
        return GREEN
    if (x == purple_color):
        return PURPLE
    if (x == cyan_color):
        return CYAN
    crd = getCoord(x1,y1)
    for i in range(76):
        for j in range(76):
            curx = crd[0]+i-38
            cury = crd[1]+j-38
            val = getColor1(im.getpixel((curx, cury)))
            if val > 0:
                return val+6
    return 0

def check_field():
    tot = 0
    for i in range(7):
        for j in range(7):
            if matr[i][j] > 0:
                tot += 1
    time.sleep(0.5)
    im = screenGrab()
    for i in range(7):
        for j in range(7):
            matr1[i][j] = getColor(im.getpixel(getCoord(i,j)), im, i, j)
            if matr1[i][j] > 6:
                matr1[i][j] -= 6
                b1[i][j] = 1
            else:
                b1[i][j] = 0
    if (tot > 46 and matr == matr1):
        return 1
    return 0

def normalize(arr):
    for i in range(7):
        a = []
        for j in range(7):
            if arr[i][j] != -1:
                a.append(arr[i][j])
        sz = len(a)
        for j in range(7-sz):
            arr[i][j] = -1
        for j in range(sz):
            arr[i][7-sz+j] = a[j]
    return arr

def check_move(arr):
    for i in range(5):
        for j in range(5):
            if arr[i][j] == arr[i+1][j] and arr[i+1][j] == arr[i+2][j] and arr[i][j] != -1 and\
                            arr[i+1][j] == arr[i+1][j+1] and arr[i+1][j+1] == arr[i+1][j+2]:
                arr[i][j] = arr[i+1][j] = arr[i+2][j] = arr[i+1][j+1] = -1
                arr = normalize(arr)
                return 8 + check_move(arr)
            if arr[i][j] == arr[i][j+1] and arr[i][j+1] == arr[i][j+2] and arr[i][j] != -1 and\
                            arr[i][j+1] == arr[i+1][j+1] and arr[i+1][j+1] == arr[i+2][j+1]:
                arr[i][j] = arr[i][j+1] = arr[i+1][j+1] = arr[i+2][j+1] = -1
                arr = normalize(arr)
                return 8 + check_move(arr)
            if arr[i+2][j] == arr[i+2][j+1] and arr[i+2][j+1] == arr[i+2][j+2] and arr[i+2][j] != -1 and\
                            arr[i+2][j+1] == arr[i+1][j+1] and arr[i+1][j+1] == arr[i][j+1]:
                arr[i+2][j] = arr[i+2][j+1] = arr[i+1][j+1] = arr[i][j+1] = -1
                arr = normalize(arr)
                return 8 + check_move(arr)
            if arr[i][j+2] == arr[i+1][j+2] and arr[i+1][j+2] == arr[i+2][j+2] and arr[i][j+2] != -1 and\
                            arr[i+1][j+2] == arr[i+1][j+1] and arr[i+1][j+1] == arr[i+1][j]:
                arr[i][j+2] = arr[i+2][j+2] = arr[i+1][j+1] = arr[i+1][j] = -1
                arr = normalize(arr)
                return 8 + check_move(arr)
            if arr[i][j] == arr[i+1][j] and arr[i+1][j] == arr[i+2][j] and arr[i][j] != -1 and\
                            arr[i][j] == arr[i][j+1] and arr[i][j+1] == arr[i][j+2]:
                arr[i][j] = arr[i+1][j] = arr[i+2][j] = arr[i][j+1] = -1
                arr = normalize(arr)
                return 8 + check_move(arr)
            if arr[i][j] == arr[i][j+1] and arr[i][j+1] == arr[i][j+2] and arr[i][j] != -1 and\
                            arr[i][j+2] == arr[i+1][j+2] and arr[i+1][j+2] == arr[i+2][j+2]:
                arr[i][j] = arr[i][j+1] = arr[i+1][j+2] = arr[i+2][j+2] = -1
                arr = normalize(arr)
                return 8 + check_move(arr)
            if arr[i][j] == arr[i+1][j] and arr[i+1][j] == arr[i+2][j] and arr[i][j] != -1 and\
                            arr[i+2][j] == arr[i+2][j+1] and arr[i+2][j+1] == arr[i+2][j+2]:
                arr[i][j] = arr[i+1][j] = arr[i+2][j] = arr[i+2][j+1] = -1
                arr = normalize(arr)
                return 8 + check_move(arr)
            if arr[i+2][j] == arr[i+2][j+1] and arr[i+2][j+1] == arr[i+2][j+2] and arr[i+2][j] != -1 and\
                            arr[i][j+2] == arr[i+1][j+2] and arr[i+1][j+2] == arr[i+2][j+2]:
                arr[i+2][j] = arr[i+2][j+1] = arr[i+1][j+2] = arr[i][j+2] = -1
                arr = normalize(arr)
                return 8 + check_move(arr)
            if arr[i+1][j+1] == arr[i+1][j] and arr[i+1][j+1] == arr[i][j+1] and arr[i+1][j] != -1 and\
                            arr[i+1][j+1] == arr[i+1][j+2] and arr[i+1][j+1] == arr[i+2][j+1]:
                arr[i+1][j+1] = arr[i+1][j] = arr[i][j+1] = arr[i+2][j+1] = -1
                arr = normalize(arr)
                return 8 + check_move(arr)
    for i in range(3):
        for j in range(7):
            if arr[i][j] == arr[i+1][j] and arr[i][j] == arr[i+2][j] and arr[i][j] != -1 and\
                            arr[i][j] == arr[i+3][j] and arr[i][j] == arr[i+4][j]:
                arr[i][j] = arr[i+1][j] = arr[i+3][j] = arr[i+4][j] = -1
                arr = normalize(arr)
                return 6 + check_move(arr)
    for i in range(7):
        for j in range(3):
            if arr[i][j] == arr[i][j+1] and arr[i][j] == arr[i][j+2] and arr[i][j] != -1 and\
                            arr[i][j] == arr[i][j+3] and arr[i][j] == arr[i][j+4]:
                arr[i][j] = arr[i][j+1] = arr[i][j+2] = arr[i][j+3] = -1
                arr = normalize(arr)
                return 6 + check_move(arr)
    for i in range(4):
        for j in range(7):
            if arr[i][j] == arr[i+1][j] and arr[i][j] == arr[i+2][j] and arr[i][j] != -1 and\
                            arr[i][j] == arr[i+3][j]:
                arr[i][j] = arr[i+1][j] = arr[i+3][j] = -1
                arr = normalize(arr)
                return 4 + check_move(arr)
    for i in range(7):
        for j in range(4):
            if arr[i][j] == arr[i][j+1] and arr[i][j] == arr[i][j+2] and arr[i][j] != -1 and\
                            arr[i][j] == arr[i][j+3]:
                arr[i][j] = arr[i][j+1] = arr[i][j+2] = -1
                arr = normalize(arr)
                return 4 + check_move(arr)
    for i in range(5):
        for j in range(7):
            if arr[i][j] == arr[i+1][j] and arr[i][j] == arr[i+2][j] and arr[i][j] != -1 and\
                                            b[i][j] + b[i+1][j] + b[i+2][j] > 0:
                arr[i][j] = arr[i+1][j] = arr[i+2][j] = -1
                arr = normalize(arr)
                return 2 + check_move(arr)
    for i in range(7):
        for j in range(5):
            if arr[i][j] == arr[i][j+1] and arr[i][j] == arr[i][j+2] and arr[i][j] != -1 and\
                                            b[i][j] + b[i][j+1] + b[i][j+2] > 0:
                arr[i][j] = arr[i][j+1] = arr[i][j+2] = -1
                arr = normalize(arr)
                return 2 + check_move(arr)
    for i in range(5):
        for j in range(7):
            if arr[i][j] == arr[i+1][j] and arr[i][j] != -1 and arr[i][j] == arr[i+2][j]:
                arr[i][j] = arr[i+1][j] = arr[i+2][j] = -1
                arr = normalize(arr)
                return 1 + check_move(arr)
    for i in range(7):
        for j in range(5):
            if arr[i][j] == arr[i][j+1] and arr[i][j] != -1 and arr[i][j] == arr[i][j+2]:
                arr[i][j] = arr[i][j+1] = arr[i][j+2] = -1
                arr = normalize(arr)
                return 1 + check_move(arr)
    return 0
def move():
    ansx1 = 0
    ansx2 = 0
    ansy1 = 0
    ansy2 = 0
    ans = 0
    arr = [0, 0, 0, 0, 0, 0, 0]
    for i in range(7):
        for j in range(7):
            arr[matr[i][j]] += 1
    for i in range(7):
        for j in range(7):
            if matr[i][j] == 0:
                xx = i+1
                yy = j
                an = 0
                if i+1 < 7 and arr[matr[i+1][j]] > an:
                    an = arr[matr[i+1][j]]
                    xx = i+1
                    yy = j
                if i-1 > 0 and arr[matr[i-1][j]] > an:
                    an = arr[matr[i-1][j]]
                    xx = i-1
                    yy = j
                if j+1 < 7 and arr[matr[i][j+1]] > an:
                    an = arr[matr[i][j+1]]
                    xx = i
                    yy = j+1
                if j-1 > 0 and arr[matr[i][j-1]] > an:
                    an = arr[matr[i][j-1]]
                    xx = i
                    yy = j-1
                makeMove(i,j,xx,yy)
                return
    for i in range(7):
        for j in range(7):
            if i+1 < 7:
                matr[i][j], matr[i+1][j] = matr[i+1][j], matr[i][j]
                b[i][j], b[i+1][j] = b[i+1][j], b[i][j]
                tmp = []
                for k in range(7):
                    tmp.append([])
                    for l in range(7):
                        tmp[k].append(matr[k][l])
                tmp = check_move(tmp)
                if tmp > ans or (tmp == ans and abs(4-ansx1) + abs(4-ansy1) > abs(4-i) + abs(4-j)):
                    ans = tmp
                    ansx1 = i
                    ansx2 = i+1
                    ansy1 = j
                    ansy2 = j
                matr[i][j], matr[i+1][j] = matr[i+1][j], matr[i][j]
                b[i][j], b[i+1][j] = b[i+1][j], b[i][j]
            if j+1 < 7:
                matr[i][j], matr[i][j+1] = matr[i][j+1], matr[i][j]
                b[i][j], b[i][j+1] = b[i][j+1], b[i][j]
                tmp = []
                for k in range(7):
                    tmp.append([])
                    for l in range(7):
                        tmp[k].append(matr[k][l])
                tmp = check_move(tmp)
                if tmp > ans or (tmp == ans and abs(4-ansx1) + abs(4-ansy1) > abs(4-i) + abs(4-j)):
                    ans = tmp
                    ansx1 = i
                    ansx2 = i
                    ansy1 = j
                    ansy2 = j+1
                matr[i][j], matr[i][j+1] = matr[i][j+1], matr[i][j]
                b[i][j], b[i][j+1] = b[i][j+1], b[i][j]
    makeMove(ansx1, ansy1, ansx2, ansy2)

def check_jew(im):
    (xw, xh) = im.size
    img = list(im.getdata())
    print(len(img))
    print(xw, xh)
    for x in range (xh-10):
        for y in range(xw-10):
            a = []
            if img[x*xw+y] != jew[0]:
                continue
            for i in range(x, x+3):
                for j in range(y, y+3):
                    a.append(img[i*xw+j])
            if a == jew:
                mousePos((y,x))
                leftDown()
                leftUp()

def take_prize():
    print('take')
    mousePos(take_action)
    leftDown()
    leftUp()
    time.sleep(10.0)

def close_jew():
    print('close')
    mousePos(close_jew_action)
    leftDown()
    leftUp()
    time.sleep(4.0)

def zoom_in():
    print('zoom-in')
    mousePos(zoom_in_action)
    leftDown()
    leftUp()
    leftDown()
    leftUp()
    time.sleep(4.0)

def add_energy_if_needed():
    im = screenGrab()
    if im.getpixel(energy_pointer) != energy_pointer_color_if_full:
        print('open store')
        mousePos(open_store_action)
        leftUp()
        leftDown()
        leftUp()
        leftDown()
        time.sleep(5.0)
        im = screenGrab()
        if im.getpixel(energy_item_action) == energy_item_color:
            print('take energy')
            mousePos(energy_item_action)
            leftUp()
            leftDown()
            leftUp()
            leftDown()
            time.sleep(4.0)
        else:
            print('close store')
            mousePos(close_store_action)
            leftUp()
            leftDown()
            leftUp()
            leftDown()
            time.sleep(4.0)

def wait_for_energy():
    im = screenGrab()
    if im.getpixel(energy_pointer) != energy_pointer_color_if_full:
        while (1):
            print('waiting for energy')
            im = screenGrab()
            if im.getpixel(energy_pointer) == energy_pointer_color_if_full:
                break
            time.sleep(120)

def launch_game():
    print('open jew')
    mousePos(open_jew_action)
    leftDown()
    leftUp()
    time.sleep(2.0)
    print('arrow')
    mousePos(arrow_action)
    leftDown()
    leftUp()
    time.sleep(2.0)
    print('start')
    mousePos(start_game_action)
    leftDown()
    leftUp()

def goGame():
    im = screenGrab()

    moves = 0
    while (1 == 1):

        moves += 1

        for i in range(7):
            for j in range(7):
                matr[i][j] = 0
                b[i][j] = 0
        while (check_field() == 0):
            im = screenGrab()
            if im.getpixel(take_pixel) == take_color:
                take_prize()
                close_jew()
                zoom_in()
                add_energy_if_needed()
                wait_for_energy()
                launch_game()

            for i in range(7):
                for j in range(7):
                    matr[i][j] = getColor(im.getpixel(getCoord(i,j)), im, i, j)
                    if matr[i][j] > 6:
                        matr[i][j] -= 6
                        b[i][j] = 1
                    else:
                        b[i][j] = 0
        move()

def main():
    goGame()
    """time.sleep(5.0)
    for i in range(500):
        leftDown()
        leftUp()
        time.sleep(0.02)"""

 
if __name__ == '__main__':
    main()