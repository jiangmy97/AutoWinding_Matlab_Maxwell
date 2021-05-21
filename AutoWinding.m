% AutoWinding.m
% For arranging windings and auto excitation arrangement in Ansys Maxwell.
% By JIANG M. Y. on May 21st, 2021.

% ===Instruction===
% 1. Fill the "input" according to your model.
% 2. Run the script in "Tools > Run script"
% 3. The name of winding terminal should be in the form of "TerminalName_TerminalNumber", (e.g. Terminal_3). 

% ===!ATTENTION!===
% There should be no variable named as "N" in your Maxwell 2D/3D > Design properties!
% For mode 1 to 4, the "Excitations" in your Maxwell design MUST BE EMPTY! 

clear all
clc

% Input 
mode=3;         
% mode=1 for single layer auto winding excitation arrangement in Maxwell 2D 
% mode=2 for double layer auto winding excitation arrangement in Maxwell 2D
% mode=3 for single layer auto winding excitation arrangement in Maxwell 3D
% mode=4 for double layer auto winding excitation arrangement in Maxwell 3D
% mode=5 for winding arrangement display only
OutputFilename='test1.vbs';        % Filename of .vbs
ProjectName='MG2-3D';         % Filename of .aedt
DesignName='noload';        % Name of the design
TerminalName1='Coil';     % Name of the winding terminal
TerminalName2='Coilu';     % Name of the winding terminal (For mode 2 and mode 4)
m=3;        % Phase
p=2;       % Pole pairs
z=12;       % Slots
coil_pitch=3;       % 0 for auto coil pitch calculation, [1, +inf) for manual control
Nc=4;        % Number of one layer conductors

% Auto calculate the coil pitch (usually 5/6 of whole pitch)
if coil_pitch==0
    coil_pitch=floor(z/(2*p)*5/6);
    if coil_pitch==0
        coil_pitch=1;
    end
end

alf=p*360/z;        % Electric angle between two slots in deg
% Display
fprintf('Coil pitch = %d slot(s)\n',coil_pitch)
fprintf('Slot angle = %.4f deg\n',alf)

% Calculate the electric angle of every slot (0 deg - 359 deg)
for i=1:z
    slot_deg(i)=(i-1)*alf;
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

% Sort the phase of slots by 60 deg band
for i=1:z
    if slot_deg(i)>=0 && slot_deg(i)<60
        slot_phase_in(i)=1;
    end
    if slot_deg(i)>=60 && slot_deg(i)<120
        slot_phase_in(i)=-3;
    end
    if slot_deg(i)>=120 && slot_deg(i)<180
        slot_phase_in(i)=2;
    end
    if slot_deg(i)>=180 && slot_deg(i)<240
        slot_phase_in(i)=-1;
    end
    if slot_deg(i)>=240 && slot_deg(i)<300
        slot_phase_in(i)=3;
    end
    if slot_deg(i)>= 300 && slot_deg(i)<360
        slot_phase_in(i)=-2;
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
for i=1:2
    for j=1:z
        switch SLOT(i,j)
            case 1
                SHOW{i,j}='+A';
            case 2
                SHOW{i,j}='+B';
            case 3
                SHOW{i,j}='+C';
            case -1
                SHOW{i,j}='-A';
            case -2
                SHOW{i,j}='-B';
            case -3
                SHOW{i,j}='-C';
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
    fprintf(fid,'oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", Array("NAME:PropServers",  _\n');
    fprintf(fid,'  "LocalVariables"), Array("NAME:NewProps", Array("NAME:N", "PropType:=", "VariableProp", "UserDef:=",  _\n');
    fprintf(fid,'  true, "Value:=", "%d"))))\n',Nc);
    % ---Add winding---
    Winding={'A','B','C'};
    for i=1:length(Winding)
        fprintf(fid,'oModule.AssignWindingGroup Array("NAME:Winding%c", "Type:=", "Current", "IsSolid:=",  _\n',Winding{i});
        fprintf(fid,'  false, "Current:=", "0mA", "Resistance:=", "0ohm", "Inductance:=", "0nH", "Voltage:=",  _\n');
        fprintf(fid,'  "0mV", "ParallelBranchesNum:=", "1")\n');
    end
    % ---Arrange winding---
    
    for i=1:length(SLOT)
        if mode==1 || mode==2
            fprintf(fid,'oModule.AssignCoil Array("NAME:%s_%d", "Objects:=", Array( _\n',TerminalName1,SLOT(3,i));
        else
            fprintf(fid,'oModule.AssignCoilTerminal Array("NAME:%s_%d", "Objects:=", Array( _\n',TerminalName1,SLOT(3,i));
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
        switch abs(SLOT(2,i))
            case 1
                fprintf(fid,'"WindingA", Array("%s_%d")\n',TerminalName1,SLOT(3,i));
            case 2
                fprintf(fid,'"WindingB", Array("%s_%d")\n',TerminalName1,SLOT(3,i));
            case 3
                fprintf(fid,'"WindingC", Array("%s_%d")\n',TerminalName1,SLOT(3,i));
        end
    end
    
    
    if mode==2 || mode==4
        for i=1:length(SLOT)
            if mode==1 || mode==2
                fprintf(fid,'oModule.AssignCoil Array("NAME:%s_%d", "Objects:=", Array( _\n',TerminalName2,SLOT(3,i));
            else
                fprintf(fid,'oModule.AssignCoilTerminal Array("NAME:%s_%d", "Objects:=", Array( _\n',TerminalName2,SLOT(3,i));
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
            switch abs(SLOT(1,i))
                case 1
                    fprintf(fid,'"WindingA", Array("%s_%d")\n',TerminalName2,SLOT(3,i));
                case 2
                    fprintf(fid,'"WindingB", Array("%s_%d")\n',TerminalName2,SLOT(3,i));
                case 3
                    fprintf(fid,'"WindingC", Array("%s_%d")\n',TerminalName2,SLOT(3,i));
            end
        end
    end
end

