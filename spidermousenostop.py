import mouse
import pyautogui
import keyboard
import time
import math
import win32api
import win32gui
import win32con
import pygetwindow
import pyautogui

max_move_speed = 1000
stopstartkey = 'PgUp'
quitkey = 'End'
catchkey = 'shift'

screen_width, screen_height = pyautogui.size()
centerx = screen_width / 2
centery = screen_height / 2
bounce = 0.95
x, y =  mouse.get_position()
CanToggle = True
Running = True
Runall = True
gettime = time.perf_counter()
xmomentum = 0
ymomentum = 0
movex = 0
movey = 0
friction = 0.98
slow = 2
pull = 0.25
gravity = 0.7
castrope = False
broken = False
fps = 60
shakewindowintensity = 0
checkedwindows = False

def is_user_grabbing_window():
    return False
    mouse_x, mouse_y = win32gui.GetCursorPos()
    active_window = pygetwindow.getActiveWindow()
    if active_window is not None and not active_window.isMinimized:  
        # Check if the active window is not the desktop window or the shell window
        if active_window.title != "Program Manager" and active_window.title != "Shell_TrayWnd":
            window_rect = active_window.topleft
            window_width, window_height = active_window.size
            if (window_rect[0] <= mouse_x <= window_rect[0] + window_width) and \
               (window_rect[1] <= mouse_y <= window_rect[1] + window_height):
                   if not keyboard.is_pressed(catchkey):
                        return False
    return False

def move_active_window(dx, dy):
    try:
        active_window = pygetwindow.getActiveWindow()
        if active_window is not None:
            new_x = active_window.left + dx
            new_y = active_window.top + dy
            active_window.moveTo(new_x, new_y)
    except Exception:
        pass 
        
def draw_rope(x1, y1, x2, y2, R, G, B):
    distance = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
    color = win32api.RGB(R, G, B)
    hdc = win32gui.GetDC(0)
    pen = win32gui.CreatePen(win32con.PS_SOLID, 1, color)
    win32gui.SelectObject(hdc, pen)
    win32gui.MoveToEx(hdc, int(x1), int(y1))
    win32gui.LineTo(hdc, int(x2), int(y2))
    win32gui.DeleteObject(pen)
    win32gui.ReleaseDC(0, hdc)
    
while Runall:
    deltaTime = (time.perf_counter() - gettime)
    gettime = time.perf_counter()
    
    if keyboard.is_pressed(quitkey):
        Runall = False
        
    if Running == True:
        if keyboard.is_pressed(stopstartkey):
            if CanToggle == True:
                Running = False
                CanToggle = False
        else:
            CanToggle = True
        
        x_prev, y_prev = x, y 
        x, y = mouse.get_position()
        
        totalxmovement = x - x_prev
        totalymovement = y - y_prev
        xmomentum += (x - x_prev - round(xmomentum))/slow + (round(movex)-round(movex)/slow)
        ymomentum += (y - y_prev - round(ymomentum))/slow + (round(movey)-round(movey)/slow) + (gravity-gravity/slow) + gravity

        if castrope:
            xmomentum *= 0.99
            ymomentum *= 0.93
        else:        
            xmomentum *= friction
            ymomentum *= friction 
            ymomentum += gravity-gravity*friction
        
        if keyboard.is_pressed(catchkey):
            if not castrope:
                tempx, tempy = mouse.get_position()
                castrope = True
                centerx = tempx + xmomentum*tempy*0.03
                centerx -= keyboard.is_pressed("left_arrow")*tempy*0.5
                centerx += keyboard.is_pressed("right_arrow")*tempy*0.5
                centery = 0
            dx = centerx - x
            dy = centery - y
            distance = math.sqrt(dx ** 2 + dy ** 2)
            if distance > 0:
                speed = pull * deltaTime
                movex = dx / distance * min(speed * distance, max_move_speed)
                movey = dy / distance * min(speed * distance, max_move_speed)
            else:
                movey = 0
                movex = 0
            draw_rope(x,y,centerx,centery,255,255,255)
        else:
            #tempx, tempy = mouse.get_position()
            #draw_rope((tempx + xmomentum*tempy*0.03)-1,0,(tempx + xmomentum*tempy*0.03)-1,15,0,255,0)
            #draw_rope((tempx + xmomentum*tempy*0.03)+1,0,(tempx + xmomentum*tempy*0.03)+1,15,255,0,0)
            if castrope:
                castrope = False
                movey = 0
                movex = 0
        
                
        if not is_user_grabbing_window():
            mouse.move(round(x + round(xmomentum) + round(movex)*castrope),round( y + round(ymomentum) + round(movex)*castrope ))
            tempx, tempy = mouse.get_position()
            if tempx >= screen_width-1 or tempx <= 0:
                xmomentum *= -bounce
                mouse.move(round(x) , round(y))
            if tempy >= screen_height-1 or tempy <= 0:
                ymomentum *= -bounce
                mouse.move(round(x) , round(y))
            else: 
                if checkedwindows:
                    checkedwindows = False
    else:
        screen_width, screen_height = pyautogui.size()
        centerx = screen_width / 2
        centery = screen_height / 2
        x, y =  mouse.get_position()
        if keyboard.is_pressed(stopstartkey):
            if CanToggle == True:
                Running = True
                CanToggle = False
        else:
            CanToggle = True
    time.sleep(1/fps)



