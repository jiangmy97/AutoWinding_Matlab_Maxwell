# AutoWinding_Matlab_Maxwell
An auto winding/excitation arrangement MATLAB script for ANSYS Maxwell 2D/3D.

The MATLAB script is [here](https://github.com/cannonjiang97/AutoWinding_Matlab_Maxwell/blob/main/AutoWinding.m).  

Last updated on May 2021-07-29.

## Function:
* Auto winding arrangement calculation. The slot-star vectogram will be plotted. The slot-phase arrangements and corresponding data will be displayed in the command window. 
* Auto excitation arrangement in ANSYS Maxwell 2D/3D. No need for manual input.

## How to use it:
1.	Fill the "input" according to your model. The items you need to fill are:
  -	Mode of this program (*mode*). The five modes of this program are:
    -	Mode 1 for single layer auto winding excitation arrangement in Maxwell 2D.
    -	Mode 2 for double layer auto winding excitation arrangement in Maxwell 2D.
    -	Mode 3 for single layer auto winding excitation arrangement in Maxwell 3D.
    -	Mode 4 for double layer auto winding excitation arrangement in Maxwell 3D.
    -	Mode 5 for winding arrangement display only.
  -	Output filename of the .vbs (*OutputFilename*).
  -	Project name of your .aedt file (*ProjectName*).
  -	Design name in your project (*DesignName*).
  -	Terminal name of your first layer coil (*TerminalName1*).
  -	Terminal name of your second layer coil for mode 2 and 4 (*TerminalName2*).
  -	Phase number of the machine (*m*).
  -	Pole-pair number of the machine (*p*).
  -	Slot number of the machine (*z*).
  -	Coil pitch of the machine (*coil_pitch*). 0 for auto coil pitch calculation. [1, +inf) for manual control.
  -	Number of one layer conductors (*N*).

2.	Run the .m file. The .vbs will generate in the same folder of .m file. The slot-star vectogram will be plotted and the slot-phase arrangements will be displayed in the command window. **Close the MATLAB after you have finished this step**.

3.	Open the .aedt file. **Make sure there are no items in “Excitations”**. **Also make sure that there should be no variable named as "N" in your Maxwell 2D/3D > Design properties.**
 
4.	Run the script in “Tools > Run script”. The excitation will automatically be arranged.
