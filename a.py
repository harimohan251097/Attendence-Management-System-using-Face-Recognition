#import module from tkinter for UI
from tkinter import *
from playsound import playsound
import os
from datetime import datetime;
import cv2
import os
import numpy as np
from PIL import Image;
import xlwt;
from datetime import datetime;
from xlrd import open_workbook;
from xlwt import Workbook;
from xlutils.copy import copy
from pathlib import Path
#creating instance of TK
root=Tk()

root.configure(background="white")

def output(filename, sheet,num, name, present):
    my_file = Path('firebase/attendance_files/'+filename+str(datetime.now().date())+'.xls');
    if my_file.is_file():
        rb = open_workbook('firebase/attendance_files/'+filename+str(datetime.now().date())+'.xls');
        book = copy(rb);
        sh = book.get_sheet(0)
        # file exists
    else:
        book = xlwt.Workbook()
        sh = book.add_sheet(sheet)
    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
                         num_format_str='#,##0.00')
    style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

    #variables = [x, y, z]
    #x_desc = 'Display'
    #y_desc = 'Dominance'
    #z_desc = 'Test'
    #desc = [x_desc, y_desc, z_desc]
    sh.write(0,0,datetime.now().date(),style1);


    col1_name = 'Name'
    col2_name = 'Present'


    sh.write(1,0,col1_name,style0);
    sh.write(1, 1, col2_name,style0);

    sh.write(num+1,0,name);
    sh.write(num+1, 1, present);
    #You may need to group the variables together
    #for n, (v_desc, v) in enumerate(zip(desc, variables)):
    fullname=filename+str(datetime.now().date())+'.xls';
    book.save('firebase/attendance_files/'+fullname)
    return fullname;

#root.geometry("300x300")
def assure_path_exists(path):
    dir = os.path.dirname(path)
    if not os.path.exists(dir):
        os.makedirs(dir)
		
def getImagesAndLabels(path):
		detector= cv2.CascadeClassifier("haarcascade_frontalface_default.xml");
		#get the path of all the files in the folder
		imagePaths=[os.path.join(path,f) for f in os.listdir(path)]
		#create empth face list
		faceSamples=[]
		#create empty ID list
		Ids=[]
		#now looping through all the image paths and loading the Ids and the images
		for imagePath in imagePaths:
			#loading the image and converting it to gray scale
			pilImage=Image.open(imagePath).convert('L')
			#Now we are converting the PIL image into numpy array
			imageNp=np.array(pilImage,'uint8')
			#getting the Id from the image
			Id=int(os.path.split(imagePath)[-1].split(".")[1])
			# extract the face from the training image sample
			faces=detector.detectMultiScale(imageNp)
			#If a face is there then append that in the list as well as Id of it
			for (x,y,w,h) in faces:
				faceSamples.append(imageNp[y:y+h,x:x+w])
				Ids.append(Id)
		return faceSamples,Ids

def function1():
    
	roll=input("enter Roll no. : ")
	name=input(" enter name : ")
	c1=""
	try:
		c=open(os.getcwd()+"/roll.txt",'r')
		c1=c.read()
		c.close()
	except:
		pass
	c=open(os.getcwd()+"/roll.txt",'w')
	if c1=="":
		c.write(roll+" ")
	else:
		c.write(c1+" "+roll)
	c.close()
	b1=""
	try:
		b=open(os.getcwd()+"/name.txt",'r')
		b1=b.read()
		b.close()
	except:
		pass
	b=open(os.getcwd()+"/name.txt",'w')
	if b1=="":
		b.write(name+" ")
	else: 
		b.write(b1+" "+name)
	b.close()
	c=open(os.getcwd()+"/roll.txt",'r')
	c1=c.read()#roll
	print(c1)
	c1=c1.split(" ")
	b=open(os.getcwd()+"/name.txt",'r')
	b1=b.read() #name
	print(b1)
	b1=b1.split(" ")
	
	b.close()
	c.close()
		
	face_id=roll
	# Start capturing video 
	vid_cam = cv2.VideoCapture(0)

	# Detect object in video stream using Haarcascade Frontal Face
	face_detector = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')

	# Initialize sample face image
	count = 0

	assure_path_exists("dataset/")

	# Start looping
	while(True):

		# Capture video frame
		_, image_frame = vid_cam.read()

		# Convert frame to grayscale
		gray = cv2.cvtColor(image_frame, cv2.COLOR_BGR2GRAY)

		# Detect frames of different sizes, list of faces rectangles
		faces = face_detector.detectMultiScale(gray, 1.3, 5)

		# Loops for each faces
		for (x,y,w,h) in faces:

			# Crop the image frame into rectangle
			cv2.rectangle(image_frame, (x,y), (x+w,y+h), (255,0,0), 2)
			
			# Increment sample face image
			count += 1

			# Save the captured image into the datasets folder
			cv2.imwrite("dataset/User." + str(face_id) + '.' + str(count) + ".jpg", gray[y:y+h,x:x+w])

			# Display the video frame, with bounded rectangle on the person's face
			cv2.imshow('frame', image_frame)

		# To stop taking video, press 'q' for at least 100ms
		if cv2.waitKey(100) & 0xFF == ord('q'):
			break

		# If image taken reach 100, stop taking video
		elif count>=30:
			print("Successfully Captured")
			break

	# Stop video
	vid_cam.release()

	# Close all started windows
	cv2.destroyAllWindows()

    
def function2():
		
		faces,Ids = getImagesAndLabels('dataSet')
		recognizer = cv2.face.LBPHFaceRecognizer_create()
		recognizer.train(faces, np.array(Ids))
		print("Successfully trained")
		b=open(os.getcwd()+"/name.txt",'r')
		b1=b.read() #name
		print(b1)
		b.close()
		recognizer.write('trainer/trainer.yml')

def function3():
    
    playsound('sound.mp3')
   
def function6():

    root.destroy()

def attend():
    os.startfile(os.getcwd()+"/firebase/attendance_files/attendance"+str(datetime.now().date())+'.xls')

#stting title for the window
root.title("AUTOMATIC ATTENDANCE MANAGEMENT USING FACE RECOGNITION")

#creating a text label
Label(root, text="FACE RECOGNITION ATTENDANCE SYSTEM",font=("times new roman",20),fg="white",bg="maroon",height=2).grid(row=0,rowspan=2,columnspan=2,sticky=N+E+W+S,padx=5,pady=5)

#creating first button
Button(root,text="Create Dataset",font=("times new roman",20),bg="#0D47A1",fg='white',command=function1).grid(row=3,columnspan=2,sticky=W+E+N+S,padx=5,pady=5)

#creating second button
Button(root,text="Train Dataset",font=("times new roman",20),bg="#0D47A1",fg='white',command=function2).grid(row=4,columnspan=2,sticky=N+E+W+S,padx=5,pady=5)

#creating third button
Button(root,text="Recognize + Attendance",font=('times new roman',20),bg="#0D47A1",fg="white",command=function3).grid(row=5,columnspan=2,sticky=N+E+W+S,padx=5,pady=5)

#creating attendance button
Button(root,text="Attendance Sheet",font=('times new roman',20),bg="#0D47A1",fg="white",command=attend).grid(row=6,columnspan=2,sticky=N+E+W+S,padx=5,pady=5)

Button(root,text="Exit",font=('times new roman',20),bg="maroon",fg="white",command=function6).grid(row=9,columnspan=2,sticky=N+E+W+S,padx=5,pady=5)


root.mainloop()
