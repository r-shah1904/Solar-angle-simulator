I was introduced with the problem of making the tracker, and given 3rd order solar equations to solve and simulate and then finally interface them in a servo. 
After a bit of research I was able to make a simulator which calculates the declination angle, azimuthal angle, solar elevation angle as well as the sunrise time, 
sunset time, and the solar noon on a particular day, I used the website www.geoastro.de to calibrate my readings according to the actual readings. I made the simulator 
in python and used the following libraries:
1. math.py
2. pandas.py
3. time.py
Further I used openpyxl to log the data into excel and also tkinter to make a GUI from this
