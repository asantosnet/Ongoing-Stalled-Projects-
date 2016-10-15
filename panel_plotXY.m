function [ColumnX,ColumnY,X,Y,posSheet] = panel_plotXY(panelSelectedFunction,Data,SheetsName,FILENAME)
%[ColumnX,ColumnY,X,Y,Data] = panel_plotXY()
%   ColumX and ColumnY sends the columns to be taken
%   X and Y the name of the values that are being used as X and Y
%   posSheet will be used later to modify Data, FILENAME and nameSheet
%
%   Input example
%   B = ones(100,5);
%     C = ones(100,6);
%     D = ones(100,5);
%     E = ones(100,5);
%     Data = { B D E;B C D};
%     panelSelectedFunction is a panel handle
%     FILENAME = {'A';'B'};
%     SheetsName = {'Aa' 'Ab';'Ba' 'Bb'};


%% WAIT - Always have this in the code - find better solution

figwait = figure('numbertitle','off','name','WAIT');
set(figwait,'Units','pixels','position',[200 200 150 150]);
set(figwait,'Resize','off');
set(figwait,'Color','red');

messsageWait = uicontrol('parent',figwait,'Style','text','Units',...
    'Normalized','position',[0.5 0.5 0.5 0.5],'String',...
    'Do not close me, When the data is saved I will disapear');



%% Back to the code

% The minimum size of the Data. 2 gives the minimul number of coluumns, the
% cellfun applies the size function to all cells and min takes the minimum
% from each column, I should probably take min(min())
[minsize,~] = min(min(cellfun('size',Data,2)));

% To show the available columns to be picked
AvailableColumns = cell(minsize,1);

for t = 1:minsize
    
    tSTR = num2str(t);
         
    AvailableColumns{t,1} = tSTR;

end

dimSheets = size(SheetsName);
posSheet = zeros(dimSheets(1),dimSheets(2));

ColumnX = '';
ColumnY = '';
X= '';
Y = '';
%% Listsbox

listboxSheet = uicontrol('parent',panelSelectedFunction,'Style','ListBox','Units',...
    'Normalized','position',[0.55,0.5,0.35,0.4],'String',...
    SheetsName(1,1:end),'Callback',@fonclistboxSheet);
% for it to allow multiselection

set(listboxSheet,'Min',0,'Max',2);

listboxFile = uicontrol('parent',panelSelectedFunction,'Style','ListBox','Units',...
    'Normalized','position',[0.03,0.5,0.35,0.4],'String',...
    FILENAME(1,1:end),'Callback',@fonclistboxFile);

%% Check Box

checkFile = uicontrol('parent',panelSelectedFunction,'Style','checkbox',...
    'Units','Normalized','position',[0.4 0.7 0.15 0.05],...
    'String','Select All Files','Callback',@fonccheckFile);

checkSheet = uicontrol('parent',panelSelectedFunction,'Style','checkbox',...
    'Units','Normalized','position',[0.4 0.6 0.15 0.05],...
    'String','Select All Sheets','Callback',@fonccheckSheet);

%% Edit text

editX = uicontrol('parent',panelSelectedFunction,'Style','edit','Units',...
                    'Normalized','position',[0.45 0.1 0.05 0.05],'String',...
                     'Vd(V)','Callback',@fonceditX);
                
editY = uicontrol('parent',panelSelectedFunction,'Style','edit','Units',...
                    'Normalized','position',[0.20 0.1 0.05 0.05],'String',...
                     'Id(A)','Callback',@fonceditY);

%% Static Text

nomSelectedX = uicontrol('parent',panelSelectedFunction,'Style','text','Units',...
    'Normalized','position',[0.3 0.2 0.15 0.1],'String',...
    'Nothing selected yet');

nomSelectedY = uicontrol('parent',panelSelectedFunction,'Style','text','Units',...
    'Normalized','position',[0.3 0.3 0.15 0.1],'String',...
    'Nothing Selected yet');

% Check if the size of the matrices are the same for the first file

if isequal_data_matrix(Data(1,:)) == 0
    textVisible = 'on'; 
else 
    textVisible = 'off';
end

mNotTS = ['All the Sheets in this file do not ' ...
           'have the same number of Columns, Attention when choosing '...
           'the Columns for X and Y since they are the same for all Sheets and files'];

nomNotTheSame = uicontrol('parent',panelSelectedFunction,'Style','text','Units',...
    'Normalized','position',[0.48 0.20 0.2 0.25],'String',...
    mNotTS ,...
    'ForegroundColor','red','Visible',textVisible);

nomSelecteX = uicontrol('parent',panelSelectedFunction,'Style','text','Units',...
    'Normalized','position',[0.03 0.2 0.15 0.1],'String',...
    'Please select X ');

nomSelecteY = uicontrol('parent',panelSelectedFunction,'Style','text','Units',...
    'Normalized','position',[0.03 0.3 0.15 0.1],'String',...
    'Please select Y ');

nomlistboxFile = uicontrol('parent',panelSelectedFunction,'Style','text',...
    'Units','Normalized','position',[0.03 0.91 0.3 0.07],...
    'String','Select the File and then choose the Sheets');

nomlistboxSheets = uicontrol('parent',panelSelectedFunction,'Style','text',...
    'Units','Normalized','position',[0.6 0.91 0.3 0.07],...
    'String','Select the Sheets to be used');

nomWhatsY = uicontrol('parent',panelSelectedFunction,'Style','text',...
    'Units','Normalized','position',[0.03 0.1 0.15 0.1],...
    'String','What does Y represents ?');

nomWhatsX = uicontrol('parent',panelSelectedFunction,'Style','text',...
    'Units','Normalized','position',[0.25 0.1 0.15 0.1],...
    'String','What does X represents ?');

% This will show the selected sheets, it is Edit so the text is in the
% center, the user cannot edit since the 'Enable' property is 'inactive',
% thus the user cannot change it and it is simpl a text 

selectedSheetsListText = uicontrol('parent',panelSelectedFunction,'Style',...
                                   'edit','String',...
                                   'Nothing yet','Units','Normalized',...
                                   'Position',[0.60 0.02 0.4 0.16],...
                                   'Enable','inactive');
     

%% Pop-up Menu

% So there is no problem in case he selects a empty sheet
if isempty(AvailableColumns) == 0
    popupXYText = AvailableColumns(1:end,1);
else
    popupXYText = 'empty';
end


popupX = uicontrol('parent',panelSelectedFunction,'Style','popupmenu','Units',...
    'Normalized','position',[0.2,0.2,0.1,0.1],'String',...
    popupXYText,'Callback',@foncpopupX);


popupY = uicontrol('parent',panelSelectedFunction,'Style','popupmenu','Units',...
    'Normalized','position',[0.2,0.3,0.1,0.1],'String',...
    popupXYText,'Callback',@foncpopupY);


%% Push Button

pushSave = uicontrol('parent',panelSelectedFunction,'style','push','Units',...
    'Normalized','position',[0.75,0.3,0.1,0.1],...
    'String','Save','Callback',@foncpushSave);

pushSelect = uicontrol('parent',panelSelectedFunction,'style','push','Units',...
    'Normalized','position',[0.92,0.65,0.06,0.1],...
    'String','Select','Callback',@foncpushSelect);

%% Callback panelSelectedFunction

%----------------------------------------------------------------------
% This function will give us the seleceted file and recover the Sheet
% name that are inside this file. It will then make the listSelectSheet
% data change in function of the selected file

    function fonclistboxFile(hObject,~)
        
        IndexSelectedFile = get(hObject,'Value');
        NewNomSheets = SheetsName(IndexSelectedFile,1:end);
        set(listboxSheet,'String',NewNomSheets);
        
        % check if the size of the matrices are the same
        if isequal_data_matrix(Data(IndexSelectedFile,:)) == 0
            textVisible = 'on';
        else
            textVisible = 'off';
        end
        
        set(nomNotTheSame,'Visible',textVisible);
        
    end

%-----------------------------------------------------------------------
% This function will allow the user to select the sheets that he wants
% to read and it creates a "shadow" matrix, i.e.  if it is not selected
% at the same position in the shadow matrix it will be 0 (false),
% othewise it will be  1 (true). OBS : posSheets has already been
% initialized as 0, so one only needs to change the true values

    function fonclistboxSheet(hObject,~)
        
        
        indexSelectedFile = get(listboxFile,'Value');
        indexSelectedSheet = get(hObject,'Value');
        
        % In case the user changes its mind and wants to change its sheets
        % selection, we reset the values to 0.
        
        posSheet(indexSelectedFile,1:end) = 0;
        
        posSheet(indexSelectedFile,indexSelectedSheet) = 1;
        
        % Write the selected positions in the text field 
        set(selectedSheetsListText,'String',mat2str(posSheet));
    end

%-------------------------------------------------------------------------
% This function saves the selected Sheets

    function foncpushSelect(~,~)
        
        indexSelectedFile = get(listboxFile,'Value');
        indexSelectedSheet = get(listboxSheet,'Value');
        
        % In case the user changes its mind and wants to change its sheets
        % selection, we reset the values to 0.
        
        posSheet(indexSelectedFile,1:end) = 0;
        
        posSheet(indexSelectedFile,indexSelectedSheet) = 1;
        
    end


%-------------------------------------------------------------------------
% This function selects all the sheets

    function fonccheckSheet(hObject,~)
        
          indexSelectedFile = get(listboxFile,'Value');
        
          % If he checks it, all sheets are selected
          
          if get(hObject,'Value') == 1
              
              posSheet(indexSelectedFile,1:end) = 1;
          else
              posSheet(indexSelectedFile,1:end) = 0;
              
              warndlg('You just unselected all the sheets, Please select them again manually');
          end
          
          % Write the selected positions in the text field 
          set(selectedSheetsListText,'String',mat2str(posSheet));
        
    end

%-------------------------------------------------------------------------
% This function selects all the files

    function fonccheckFile(hObject,~)
            
          % If he checks it, all sheets are selected
          
          if get(hObject,'Value') == 1
              
              posSheet(1:end,1:end) = 1; 
          else
              posSheet(1:end,1:end) = 0;
              warndlg('You just unselected all the files, Please select themaagain manually');
          end
          
          % Write the selected positions in the text field 
          set(selectedSheetsListText,'String',mat2str(posSheet));
        
    end

%-------------------------------------------------------------------------
% This function sleects the Columns for the X values

    function foncpopupX(hObject,~)
        
        if isempty(AvailableColumns) == 0
            XColumn =AvailableColumns{get(hObject,'Value'),1};
            set(nomSelectedX,'String',XColumn);
            ColumnX = str2num(XColumn);
            
        else
            ColumnX = '';
        end

        
    end

%-------------------------------------------------------------------------
% This function sleects the Columns for the Y values

    function foncpopupY(hObject,~)
        
        if isempty(AvailableColumns) == 0
            YColumn = AvailableColumns{get(hObject,'Value'),1};
            set(nomSelectedY,'String',YColumn);
            ColumnY = str2num(YColumn);
        else
            ColumnY = '';
        end

    end

%-------------------------------------------------------------------------
% This function gets the meaning of the X value

    function fonceditX(hObject,~)
        
        X = get(hObject,'String');

    end

%-------------------------------------------------------------------------
% This function gets the meaning of the y value

    function fonceditY(hObject,~)
        
        Y = get(hObject,'String');
    end
%-------------------------------------------------------------------------
% This function Saves everything

    function foncpushSave(~,~)
        
        % Check if the user haven't forgot anything
        
        % If he doesn't writes anythign I assume he agrees with the example
        X = get(editX,'String');
        Y = get(editY,'String');

        if isempty(ColumnX) == 0 && isempty(ColumnY) ==0 &&...
                isempty(X) == 0 && isempty(Y) ==0
            
            
            % In case no Sheets have been selected
            
            if isequal(posSheet,zeros(dimSheets(1),dimSheets(2))) == 1
                
                dlgTitle = 'User Question';
                dlgQuestion = ['You haven''t selected any Sheet',...
                    ' Do you wish to continue?'];
                choice = questdlg(dlgQuestion,dlgTitle,'Yes','No','Yes');
                
                if isequal(choice,'Yes') == 1
 
                   warndlg([ 'Saved - Do not Press Save Again / '...
                            'Remove the function from the list and add it again'...
                            'in case you need to change something']);
                        
                   uiresume(figwait);
                   close(figwait);
                end
            else
                       
               warndlg([ 'Saved - Do not Press Save Again / '...
                        'Remove the function from the list and add it again'...
                        'in case you need to change something']);
                   
                
                uiresume(figwait);
                close(figwait);
                
            end
            
        else
            
            % In case some sheets are empty, they can be ignored since they
            % will be dealt with later on, also, cells needs to have rows
            % with the same number of columns and vice versa, this makes it
            % hard to remove one sheet without leaving an empty cell.
            if isempty(AvailableColumns)
                
                
                dlgTitle = 'User Question';
                dlgQuestion = ['Some sheets are empty',...
                    ' Do you wish to ignore them? (if No choose another function'];
                choice = questdlg(dlgQuestion,dlgTitle,'Yes','No','Yes');
                
                % To allow the user to choose the columns again
                if isequal(choice,'Yes') == 1
                    
                    % Finding the index for which the sheets are zero
                    
                    nColumnsSheets = cellfun('size',Data,2);
                    nNonZeroIndex = find(nColumnsSheets);
                    
                    if isempty(nNonZeroIndex) == 0
                        
                        warndlg('Select the X and Y cloumsn again');
                        
                        % To show the available columns to be picked
                        
                        [minsize,~] = min(min(nColumnsSheets(nNonZeroIndex)));
                        AvailableColumns = cell(minsize,1);
                        
                        for tt = 1:minsize
                            
                            tSTR = num2str(tt);
                            
                            AvailableColumns{tt,1} = tSTR;
                            
                        end
                        
                        set(popupX,'String',AvailableColumns(1:end,1));
                        set(popupY,'String',AvailableColumns(1:end,1));
                        
                             
                    else
                        
                        warndlg(['All Sheets are empty, the function will' ...
                            'work, but nothing will be plotted']);
                        
                        uiresume(figwait);
                        close(figwait);
                        
                    end
                    
                end
            end
            
            warndlg('You migth forgot something / Ignore this if empty sheets warning');
        end
        
    end

uiwait(figwait);
end



function [isItEqual] = isequal_data_matrix(DataRow)
%Just checking if the Data in the Sheets have the same size
% isItEqual is a logical 1 or 0 and DataRow is a one Row cell

isItEqual = 0;

% Remove the empty cells wich migth give us an logical erro
DataRow(1,~cellfun('isempty',DataRow));

if size(DataRow,2) == 1
    
    isItEqual = 1;
    
else
   
    % Otherwise DataRow{1,:} can be only one input and isequal needs at
    % least 2
    isItEqual = isequal(DataRow{1,:});

end

end


