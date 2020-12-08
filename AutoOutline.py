import win32gui
import win32com.client
import re
import keyboard
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

from tkinter import *
from duelMonitor import *

def main():
	
	root = createMainWindow()
	
	#rulerData = getRulerData()
	#sendMinecraftMessage(rulerData)
	
	root.mainloop()

	
def getRulerData():
	
	#Find the ruler window from google earth
	rulerWindow = win32gui.FindWindow(None, 'Ruler')
	
	if not rulerWindow:
		print("Unable to find ruler window")
		return null
	
	box = win32gui.GetWindowRect(rulerWindow)
	bitmap = get_screen_buffer(box)
	img = make_image_from_buffer(bitmap)

	#Make the image larger so the characters can be detected better
	width, height = img.size
	newsize = ((width*5), (height*5))
	img = img.resize(newsize)
	img.save('enhanced.png')

	#Convert the image into readable data to use
	text = pytesseract.image_to_string(img)

	#Get the map length value from the ruler image
	ml1 = re.search(r"(Map Length:.+\d)", text)
	mt1 = ml1.group()
	ml2 = re.search(r"[^Map Length:]+", mt1)
	mt2 = ml2.group()
	mt3 = mt2.replace(',', '')
	length = mt3

	#Get the heading value from the ruler image
	h1 = re.search(r"(Heading:.+\d)", text)
	ht1 = h1.group()
	h2 = re.search(r"[^Heading:]+", ht1)
	ht2 = h2.group()
	heading = ht2
	
	#Put the length and heading into a single string and return it
	rulerdata = length + " " + heading
	return rulerdata

def sendMinecraftMessage(rulerData, pointType, origin):
	
	#Find the minecraft window
	minecraftWindow = win32gui.FindWindow("GLFW30", None)
	win32gui.SetForegroundWindow(minecraftWindow)
	minecraft = win32gui.GetFocus()
	
	minMessage = '/' + pointType + ' ' + origin + ' ' + rulerData
	
	shell = win32com.client.Dispatch('WScript.Shell')
	
	for x in minMessage:
		shell.SendKeys(x)
		
	keyboard.press_and_release('enter')
	keyboard.press_and_release('t')

def createMainWindow():

	root = Tk()
	root.title('Auto Outline Tool')
	root.geometry("356x300")
	root.configure(bg = "#5DADE2")
	
	cmdSetupFrame = Frame(root,
		width=20,
		height=10,
		bg = "#5DADE2"
		)
	
	originFrame = Frame(cmdSetupFrame,
		width=20,
		height=10,
		bg = "#5DADE2"
		)
	
	pointTypeFrame = Frame(cmdSetupFrame,
		width=20,
		height=10,
		bg = "#5DADE2"
		)
		
	connectPointsFrame = Frame(root,
		width=20,
		height=10,
		bg = "#5DADE2"
		)
		
	rFrame = Frame(connectPointsFrame,
		width=20,
		height=10,
		bg = "#5DADE2"
		)
	
	cmdSetup = Label(root,
		text ="Command Setup",
		width = 50,
		height = 2,
		bg = "#5DADE2"
		)
		
	originPoint = Label(originFrame,
		text = "Origin Point",
		width = 20,
		height = 2,
		bg = "#5DADE2"
		)
		
	originEntry = Entry(originFrame,
		width = 20,
		bg = "White"
		)
		
	pointType = Label(pointTypeFrame,
		text = "Point type",
		width = 20,
		height = 2,
		bg = "#5DADE2"
		)
	
	clicked = StringVar()
	clicked.set("Point")
	pointT = clicked.get()
	
	pointDrop = OptionMenu(pointTypeFrame,
		clicked,
		"Point", 
		"Point15",
		"Point2"
		)
	pointDrop.configure(bg = "White")
	
	connectPoints = Label(connectPointsFrame,
		text = "Connect Points",
		width = 50,
		height = 2,
		bg = "#5DADE2"
		)
	
	r = BooleanVar()
	r.set("No")
		
	r1 = Radiobutton(rFrame, 
		text="Yes", 
		variable=r, 
		value=True, 
		command=lambda: selected(r.get()),
		bg = "#5DADE2"
		)
	r2 = Radiobutton(rFrame, 
		text="No", 
		variable=r, 
		value=False, 
		command=lambda: selected(r.get()),
		bg = "#5DADE2"
		)
		
	calculateBtn = Button(root,
		text="Calculate point",
		width = 25,
		height = 2,
		bg = "White",
		command=lambda: calculatePoint(pointT, originEntry.get())
		)
		
	cmdSetup.grid(row=0,column=1)
	
	cmdSetupFrame.grid(row=1,column=1)
	
	pointTypeFrame.grid(row=0,column=0)
	pointType.pack()
	pointDrop.pack()
	
	originFrame.grid(row=0,column=1)
	originPoint.pack()
	originEntry.pack()
	
	connectPointsFrame.grid(row=2,column=1)
	connectPoints.pack()
	rFrame.pack()
	r1.grid(row=0,column=0)
	r2.grid(row=0,column=1)
	calculateBtn.grid(row=3,column=1)
	
	return root
	
def calculatePoint(pointType, origin):
	rulerData = getRulerData()
	sendMinecraftMessage(rulerData, pointType, origin)
	
	
def selected(value):
	#radio button code here
	line = value
	
main()