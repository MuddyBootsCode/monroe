from tkinter import *

root = Tk()

top_frame = Frame(root)
top_frame.pack()
bottom_frame = Frame(root)
bottom_frame.pack(side=BOTTOM)

button1 = Button(top_frame, text="Button1", fg='red')
button2 = Button(top_frame, text="Button2", fg='blue')
button3 = Button(top_frame, text="Button3", fg='yellow')
button4 = Button(bottom_frame
                 , text="Button4", fg='green')

button1.pack(side=LEFT)
button2.pack(side=LEFT)
button3.pack(side=LEFT)
button4.pack(side=LEFT)

root.mainloop()