# AutoWinding_Matlab_Maxwell
An auto AC and DC machine's winding/excitation arrangement MATLAB script for ANSYS Maxwell 2D/3D.

The MATLAB script for AC machine is [here](https://github.com/jiangmy97/AutoWinding_Matlab_Maxwell/blob/main/AutoWindingAC.m) and for DC machine is [here](https://github.com/jiangmy97/AutoWinding_Matlab_Maxwell/blob/main/AutoWindingDC.m).  

Last updated on 2021-11-2.

## Function
* Auto AC/DC machines' winding arrangement calculation. The slot-star vectogram will be plotted. The slot-phase arrangements and corresponding data will be displayed in the command window. 
* Auto excitation arrangement in ANSYS Maxwell 2D/3D. No need for manual input.

## How to use it
1.	Open the .aedt file. Use "Draw > Duplicate around axis" to duplicate your coil object, input the following parameters in the window:

Item | Value/Selection
------------ | -------------
Angle | 360deg/Number of slots (e.g. 15deg for 24 slots)
Total number | Number of slots
Attach To Original Object | Yes

2.  Right click the object, use "Edit > Boolean > Seperate Bodies" to seperate them.
3.	Open the .m file. Fill the "Input" section. Please follow the instructions behind each parameters.
4.	Run the .m file. The .vbs will generate in the same folder of .m file. The slot-star vectogram will be plotted and the slot-phase arrangements will be displayed in the command window.
5.	Open the .aedt file. Run the script in “Tools > Run script”. The excitation will automatically be arranged.
