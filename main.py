# -*- coding: utf-8 -*-
"""
Created on Wed Jul 17 12:00:47 2019

@author: MANOJ
"""

import tkinter as tk
from tkinter import *
import pyttsx3
from PIL import Image, ImageTk
import os
import docx

from tkinter import filedialog

engine = pyttsx3.init()

class Widget():
    def __init__(self):
        self.root = tk.Tk()
        self.v = tk.IntVar()
        self.flag = 0

        self.root.configure(bg = '#448aff')
        
        self.root.title('Text To Speech')

        sp = os.getcwd()
        imgicon = PhotoImage(file=os.path.join(sp,'logo.png')) #Put a logo for the title bar 
        self.root.tk.call('wm', 'iconphoto', self.root._w, imgicon)  
        
        self.label1 = tk.Label(self.root , text = 'ENTER TEXT IN THE TEXT BOX BELOW' , width = 37 , bg = '#fff9c4' , font = 'Times')
        self.label1.pack(pady = 25)
        
        self.t = Text(self.root, width=40, height=10)
        self.t.pack(padx = 160 , pady = 10)

        self.button3 = tk.Button(
            self.root,
            text='Clear Text',
            command=self.clearText,
            width=40,
            bg = '#fff9c4'
            )
        self.button3.pack(pady = 20)

        self.label2 = tk.Label(self.root , text = 'CHOOSE A VOICE' , bg = '#448aff')
        self.label2.pack()
        self.rd1 = tk.Radiobutton(
            self.root,
            text="Male US-EN",
            indicatoron = 1,
            variable=self.v,
            value=0,
            command=self.chooseMale,
            bg='#d1c4e9'
            ).pack()
        self.rd1 = tk.Radiobutton(
            self.root,
            text="Female US-EN",
            indicatoron = 1,
            variable=self.v,
            value=1,
            command=self.chooseFemale,
            bg='#d1c4e9'
            ).pack()

      
        self.button = tk.Button(
            self.root ,
            text = 'Speak' ,
            command=self.clicked ,
            width = 75 ,
            bg = '#fff9c4',
            activebackground = '#81c784'
            )
        self.button.pack(pady = 20)

        self.button2 = tk.Button(
                self.root,
                text='Choose a file',
                command=self.chooseFile,
                width=75,
                bg = '#fff9c4',
                activebackground = '#81c784'
                )
        self.button2.pack(pady = 20)
        img = ImageTk.PhotoImage(Image.open('speaker.png'))
        self.label2 = Label(self.root, image = img , width = 50 , height = 50)
        self.label2.image = img
        self.label2.pack(pady = 20)

    def clearText(self):
        self.t.delete(1.0 , END),

    def chooseMale(self):
        engine.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_DAVID_11.0')

    def chooseFemale(self):
        engine.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0')
    def speak(self , text):
        if self.flag == 1:
             engine.setProperty('rate' , 185) #Default rate : 200
             engine.say(text)
             engine.runAndWait()

        else:
            engine.say(text)
            engine.runAndWait()

        
    def clicked(self):
        text = self.t.get('1.0', END)
        self.speak(text)
        
    def speakFromFile(self , loc):

        doc = docx.Document(loc)
        fullText = []
        for para in doc.paragraphs:
            fullText.append(para.text)

        text = ""
        
        text = text.join(fullText)
        text.split('.')

        self.flag = 1
        self.t.insert(INSERT , text)
           
    
    def chooseFile(self):
        currdir = os.getcwd() 
        tempfile = filedialog.askopenfilename(parent=self.root , initialdir=currdir, title='Please select a file')
        if len(tempfile) > 0:
            print("You chose %s" % tempfile)
        fileLoc = tempfile
        self.speakFromFile(fileLoc)
        
        
        
        
if __name__ == '__main__':
    widget = Widget()
        
