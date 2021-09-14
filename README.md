# AutoWinding_Matlab_Maxwell
An auto AC and DC machine's winding/excitation arrangement MATLAB script for ANSYS Maxwell 2D/3D.

The MATLAB script for AC machine is [here](https://github.com/jiangmy97/AutoWinding_Matlab_Maxwell/blob/main/AutoWindingAC.m) and for DC machine is [here](https://github.com/jiangmy97/AutoWinding_Matlab_Maxwell/blob/main/AutoWindingDC.m).  

Last updated on 2021-08-04.

## Function
* Auto AC/DC machines' winding arrangement calculation. The slot-star vectogram will be plotted. The slot-phase arrangements and corresponding data will be displayed in the command window. 
* Auto excitation arrangement in ANSYS Maxwell 2D/3D. No need for manual input.

## How to use it
1.	Fill the "Input" section. Please follow the instructions behind each parameters.

2.	Run the .m file. The .vbs will generate in the same folder of .m file. The slot-star vectogram will be plotted and the slot-phase arrangements will be displayed in the command window.

3.	Open the .aedt file.
 
4.	Run the script in “Tools > Run script”. The excitation will automatically be arranged.

## Notice
* The terminal name are in the form of "TerminalName_TerminalNumber" (e.g. "LapCoil_2"). However, **the initial terminal doesn't have terminal number. (e.g. "LapCoil" itself is the name of the first coil)** The terminals are recommended to be created by duplicating around axis, as the program is written based on this method. E.g., if the initial terminal has the name of "LapCoil", then the terminal after duplication should be like "LapCoil_1", "LapCoil_2", etc. If the initial terminal has the name of "LapCoil1", then the terminal after duplication should be like "LapCoil1_1", "LapCoil1_2", etc.  
