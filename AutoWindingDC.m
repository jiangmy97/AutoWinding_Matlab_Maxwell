% AutoWindingDC.m
% For arranging DC machine's windings and auto excitation arrangement in Ansys Maxwell.
% By JIANG M. Y. on 2021-08-04.

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
OutputFilename='testDC.vbs';        % Filename of .vbs
ProjectName='DCPM-1';         % Filename of .aedt
DesignName='noload';        % Name of the design
TerminalName1='Coil';     % Name of the winding terminal
TerminalName2='Coilu';     % Name of the winding terminal (For mode 2 and mode 4)

p=5;       % Pole pairs
z=11;       % Slots

Nc=0;        % Number of one layer conductors, Nc=0 for already set
FlagCreateWinding=1;    % Switch of create winding, 1 for need to create, 0 for no need
%=========================================================
alf=p*360/z;        % Electric angle between two slots in deg
y1=floor(z/(2*p)); % Length of the winding in slots

fprintf('Pole-pair number= %d\n',p)
fprintf('Slots= %d\n',z)
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

% Judge winding type
z1=z/gcd(p,z);
p1=p/gcd(p,z);
if mod(z1,2)==0
    FlagWindingType=1; 
    fprintf('Winding type: Lap coil\n')
else
    FlagWindingType=2; 
    fprintf('Winding type: Wave coil\n')
end



% Arrange slots to windings
    for i=1:z1
        for j=1:gcd(p,z)
            WindingSlot(i,1+(j-1)*2)=i+(j-1)*z1;
            WindingSlot(i,2+(j-1)*2)=-(i+(j-1)*z1+y1);
        end
    end
[row,col]=size(WindingSlot);
for i=1:row
    for j=1:col
        if WindingSlot(i,j)>z
            WindingSlot(i,j)=WindingSlot(i,j)-z;
        elseif WindingSlot(i,j)<-z
            WindingSlot(i,j)=WindingSlot(i,j)+z;
        end
    end
end
% Arrange windings to slots
WindingSlot=ceil(WindingSlot);
[row, col]=size(WindingSlot);
for i=1:row
    for j=1:col
        if WindingSlot(i,j)<0
            SLOT(1,-WindingSlot(i,j))=-i;
        else
            SLOT(2,WindingSlot(i,j))=i;
        end
    end
end

for i=1:z
    SLOT(3,i)=i-1;
end

% Winding connection
if FlagWindingType==1
    for i=1:z1
        CONNECT(i)=i;
    end
elseif FlagWindingType==2
    y2=(z1-1)/p1;
    for i=1:z1
        CONNECT(i)=1+(i-1)*y2;
        while CONNECT(i)>z
            CONNECT(i)=CONNECT(i)-z;
        end
    end
end
fprintf('Winding connection:\n')
for i=1:length(CONNECT)
    fprintf('%d-',CONNECT(i))
end
fprintf('1 (loop)\n')
fprintf('* For winding-slot arrangement, please check "WindingSlot".\n')

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
        for i=1:z/p
            fprintf(fid,'oModule.AssignWindingGroup Array("NAME:Winding%d", "Type:=", "External", "IsSolid:=",  _\n',i);
            fprintf(fid,'  false, "Current:=", "0mA", "Resistance:=", "0ohm", "Inductance:=", "0nH", "Voltage:=",  _\n');
            fprintf(fid,'  "0mV", "ParallelBranchesNum:=", "1")\n');
        end
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
            fprintf(fid,'"Winding%d", Array("%s_%d")\n',abs(SLOT(2,i)),TerminalName1,SLOT(3,i));
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
            fprintf(fid,'"Winding%d", Array("%s_%d")\n',abs(SLOT(1,i)),TerminalName2,SLOT(3,i));
        end
    end
    fclose('all');
end
