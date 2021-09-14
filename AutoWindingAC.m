% AutoWinding.m
% For arranging AC machine's windings and auto excitation arrangement in Ansys Maxwell.
% By JIANG M. Y. on 2021-09-08.

% ===Instruction===
% 1. Fill the "input" according to your model.
% 2. Run the script in "Tools > Run script"
% 3. The name of winding terminal should be in the form of "TerminalName_TerminalNumber", (e.g. Terminal_3).

clear all
clc

% Input
mode=2;
% mode=1 for single layer auto winding excitation arrangement in Maxwell 2D
% mode=2 for double layer auto winding excitation arrangement in Maxwell 2D
% mode=3 for single layer auto winding excitation arrangement in Maxwell 3D
% mode=4 for double layer auto winding excitation arrangement in Maxwell 3D
% mode=5 for winding arrangement display only
OutputFilename='test1.vbs';        % Filename of .vbs
ProjectName='Project1';         % Filename of .aedt
DesignName='MaxwellDesign1';        % Name of the design
TerminalName1='Coil';     % Name of the winding terminal
TerminalName2='Coilu';     % Name of the winding terminal (For mode 2 and mode 4)
WindingNames={'A','B','C'};         % Names of windings, like 'A', 'B', 'C'
p=2;       % Pole pairs
z=12;       % Slots
coil_pitch=0;       % 0 for auto coil pitch calculation, [1, +inf) for manual control, -1 for drum coil
Nc=1;        % Number of one layer conductors, Nc=0 for already set
FlagCreateWinding=1;    % Switch of create winding, 1 for need to create, 0 for no need
%=========================================================
m= length(WindingNames); % Phases
% Auto calculate the coil pitch (usually 5/6 of whole pitch)
if coil_pitch==0
    coil_pitch=floor(z/(2*p)*5/6);
    if coil_pitch==0
        coil_pitch=1;
    end
elseif coil_pitch==-1
    coil_pitch=0;
end

alf=p*360/z;        % Electric angle between two slots in deg
% Display
fprintf('Coil pitch = %d slot(s)\n',coil_pitch)
fprintf('Slot angle = %.4f deg\n',alf)

% Calculate the electric angle of every slot (0 deg - 359 deg)
for i=1:z
    slot_deg(i)=(i-1)*alf+0.0001;
    while(1)
        if slot_deg(i)>=360
            slot_deg(i)=slot_deg(i)-360;
        else
            break
        end
    end
end

% Plot slot-star vectogram
n = z / gcd(z,p);
if n == p
    n = 2*n;
end
for i=1:n
    c(i)=complex(cosd(slot_deg(i)),sind(slot_deg(i)));
end
compass(c)
for i=1:n
    text(1.25*cosd(slot_deg(i)),1.25*sind(slot_deg(i)),num2str(i),'FontSize',18);
end

% Sort the phase of slots by band
BandDeg=360/(2*m); % Degree of one band
for i=1:m
    BandPos(i,1)=(2*i-2)*BandDeg;
    BandPos(i,2)=(2*i-1)*BandDeg;
    BandNeg(i,1)=180+(2*i-2)*BandDeg;
    BandNeg(i,2)=180+(2*i-1)*BandDeg;
end

for i=1:m
    for j=1:2
        if BandPos(i,j)>360
            BandPos(i,j)=BandPos(i,j)-360;
        end
        if BandNeg(i,j)>360
            BandNeg(i,j)=BandNeg(i,j)-360;
        end
    end
end

for i=1:z
    for j=1:length(BandPos)
        if slot_deg(i)>=BandPos(j,1) && slot_deg(i)<BandPos(j,2)
            slot_phase_in(i)=j;
        end
    end
    for j=1:length(BandNeg)
        if slot_deg(i)>=BandNeg(j,1) && slot_deg(i)<BandNeg(j,2)
            slot_phase_in(i)=-j;
        end
    end
end

% Calculate the outside layer of slots (for double layer lapcoil)
for i=1:coil_pitch
    slot_phase_out(i)=-slot_phase_in(z-coil_pitch+i);
end
for i=1:z-coil_pitch
    slot_phase_out(i+coil_pitch)=-slot_phase_in(i);
end

SLOT=[slot_phase_out;slot_phase_in];

% Calculate the pitch factor
if coil_pitch==0
    k_p=1;
else
    k_p=sind(coil_pitch/(z/(2*p))*90);
end

% Calculate the distribution factor
q=z/(2*p*m);    % Slots per phase per pole
[N,D]=rat(q);
if D~=1
    q=N;
    alf=60/N;
end
k_d=sind(q*alf/2)/(q*sind(alf/2));

% Calculate the winding factor
k_w=k_d*k_p;

% Display
fprintf('Pitch factor = %.4f\n',k_p)
fprintf('Distribution factor = %.4f\n',k_d)
fprintf('Winding factor = %.4f\n',k_w)
fprintf('++++++++++++++++++++++\n')

% Rename and print the phase of each slot
for i=1:length(WindingNames)
    WindingNamesPos{i}=['+',WindingNames{i}];
    WindingNamesNeg{i}=['-',WindingNames{i}];
end
for i=1:2
    for j=1:z
        if SLOT(i,j)>0
            SHOW{i,j}=WindingNamesPos{SLOT(i,j)};
        else
            SHOW{i,j}=WindingNamesNeg{-SLOT(i,j)};
        end
        fprintf('%s ',SHOW{i,j})
    end
    fprintf('\n')
end

for i=1:z
    SHOW{3,i}=num2str(i-1);
    SLOT(3,i)=i-1;
end

% .vbs output
if mode~=5
    % Configure output file
    Dat=fopen(OutputFilename,'w');
    fprintf(Dat,'''------------------------------------------\n');
    fprintf(Dat,'''AutoWinding by JIANG M. Y.\n');
    fprintf(Dat,'''Last Modified: %s\n',datestr(now,31));
    fprintf(Dat,'''------------------------------------------\n');
    fclose(Dat);
    fid=fopen(OutputFilename,'a');
    % ---Prefix---
    fprintf(fid,'Dim oAnsoftApp\n');
    fprintf(fid,'Dim oDesktop \n');
    fprintf(fid,'Dim oProject \n');
    fprintf(fid,'Dim oDesign \n');
    fprintf(fid,'Dim oEditor\n');
    fprintf(fid,'Dim oModule\n');
    fprintf(fid,'Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")\n');
    fprintf(fid,'Set oDesktop = oAnsoftApp.GetAppDesktop()\n');
    fprintf(fid,'oDesktop.RestoreWindow\n');
    % ---Set---
    fprintf(fid,'Set oProject = oDesktop.SetActiveProject("%s")\n',ProjectName);
    fprintf(fid,'Set oDesign = oProject.SetActiveDesign("%s")\n',DesignName);
    fprintf(fid,'Set oModule = oDesign.GetModule("BoundarySetup")\n');
    % ---Add the number of conducters per layer
    if Nc~=0
        fprintf(fid,'oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _\n');
        fprintf(fid,'  "LocalVariables"), Array("NAME:NewProps", Array("NAME:N", "PropType:=", "VariableProp", "UserDef:=",  _\n');
        fprintf(fid,'  true, "Value:=", "%d"))))\n',Nc);
    end
    % ---Add winding---
    if FlagCreateWinding==1
        for i=1:length(WindingNames)
            fprintf(fid,'oModule.AssignWindingGroup Array("NAME:Winding%s", "Type:=", "Current", "IsSolid:=",  _\n',WindingNames{i});
            fprintf(fid,'  false, "Current:=", "0mA", "Resistance:=", "0ohm", "Inductance:=", "0nH", "Voltage:=",  _\n');
            fprintf(fid,'  "0mV", "ParallelBranchesNum:=", "1")\n');
        end
    end
    % ---Arrange winding---
    
    for i=1:length(SLOT)
        if mode==1 || mode==2
            if SLOT(3,i)==0
                fprintf(fid,'oModule.AssignCoil Array("NAME:%s", "Objects:=", Array( _\n',TerminalName1);
            else
                fprintf(fid,'oModule.AssignCoil Array("NAME:%s_%d", "Objects:=", Array( _\n',TerminalName1,SLOT(3,i));
            end
        else
            if SLOT(3,i)==0
                fprintf(fid,'oModule.AssignCoilTerminal Array("NAME:%s", "Objects:=", Array( _\n',TerminalName1);
            else
                fprintf(fid,'oModule.AssignCoilTerminal Array("NAME:%s_%d", "Objects:=", Array( _\n',TerminalName1,SLOT(3,i));
            end
        end
        
        if i==1
            fprintf(fid,'  "%s"), "Conductor number:=", "N", ',TerminalName1);
        else
            fprintf(fid,'  "%s_%d"), "Conductor number:=", "N", ',TerminalName1,SLOT(3,i));
        end
        if SLOT(2,i)>0
            if mode==1 || mode==2
                fprintf(fid,'"PolarityType:=", "Positive")\n');
            else
                fprintf(fid,'"Point out of terminal:=", false)\n');
            end
        else
            if mode==1 || mode==2
                fprintf(fid,'"PolarityType:=", "Negative")\n');
            else
                fprintf(fid,'"Point out of terminal:=", true)\n');
            end
        end
        if mode==1 || mode==2
            fprintf(fid,'oModule.AddWindingCoils ');
        else
            fprintf(fid,'oModule.AddWindingTerminals ');
        end
        if SLOT(3,i)==0
            fprintf(fid,'"Winding%s", Array("%s")\n',WindingNames{abs(SLOT(2,i))},TerminalName1);
        else
            fprintf(fid,'"Winding%s", Array("%s_%d")\n',WindingNames{abs(SLOT(2,i))},TerminalName1,SLOT(3,i));
        end
    end
    
    
    if mode==2 || mode==4
        for i=1:length(SLOT)
            if mode==1 || mode==2
                if SLOT(3,i)==0
                    fprintf(fid,'oModule.AssignCoil Array("NAME:%s", "Objects:=", Array( _\n',TerminalName2);
                else
                    fprintf(fid,'oModule.AssignCoil Array("NAME:%s_%d", "Objects:=", Array( _\n',TerminalName2,SLOT(3,i));
                end
            else
                if SLOT(3,i)==0
                    fprintf(fid,'oModule.AssignCoilTerminal Array("NAME:%s", "Objects:=", Array( _\n',TerminalName2);
                else
                    fprintf(fid,'oModule.AssignCoilTerminal Array("NAME:%s_%d", "Objects:=", Array( _\n',TerminalName2,SLOT(3,i));
                end
            end
            
            if i==1
                fprintf(fid,'  "%s"), "Conductor number:=", "N", ',TerminalName2);
            else
                fprintf(fid,'  "%s_%d"), "Conductor number:=", "N", ',TerminalName2,SLOT(3,i));
            end
            if SLOT(1,i)>0
                if mode==2
                    fprintf(fid,'"PolarityType:=", "Positive")\n');
                else
                    fprintf(fid,'"Point out of terminal:=", false)\n');
                end
            else
                if mode==2
                    fprintf(fid,'"PolarityType:=", "Negative")\n');
                else
                    fprintf(fid,'"Point out of terminal:=", true)\n');
                end
            end
            if mode==1 || mode==2
                fprintf(fid,'oModule.AddWindingCoils ');
            else
                fprintf(fid,'oModule.AddWindingTerminals ');
            end
            if SLOT(3,i)==0
                fprintf(fid,'"Winding%s", Array("%s")\n',WindingNames{abs(SLOT(1,i))},TerminalName2);
            else
                fprintf(fid,'"Winding%s", Array("%s_%d")\n',WindingNames{abs(SLOT(1,i))},TerminalName2,SLOT(3,i));
            end
        end
    end
    fclose('all');
end

