function [name,location,posSheet,Close ] = selectExcelSheet(Sheets,FILENAME)
%UNTITLED Summary of this function goes here
%   Detailed explanation goes here

Close = 0;
dimSheets = size(Sheets);
posSheet = zeros(dimSheets(1),dimSheets(2));
name = '';
location = '';
% The figure
% By doing this I am storing all the data in the figmain variable ( same
% role as handle maybe?

figmain = figure('numbertitle','off','name','SelectData 1');
set(figmain,'Units','pixels','position',[200 200 737 550]);
set(figmain,'Resize','off');
set(figmain,'Color','black');


%% Push

pushNext = uicontrol('parent',figmain,'style','push','Units',...
                             'Normalized','position',[0.8,0.05,0.15,0.1],...
                             'String','Next','Callback',@foncpushNext);
                         
pushCancel = uicontrol('parent',figmain,'style','push','Units',...
                     'Normalized','position',[0.6,0.05,0.15,0.1],...
                     'String','Cancel','Callback',@foncpushCancel);
               
pushSelect = uicontrol('parent',figmain,'style','push','Units',...
                     'Normalized','position',[0.50,0.2,0.09,0.09],...
                     'String','Select','Callback',@foncpushSelect);
                 
pushUnselect = uicontrol('parent',figmain,'style','push','Units',...
                     'Normalized','position',[0.405,0.2,0.09,0.09],...
                     'String','Unselect','Callback',@foncpushUnselect);
                 
%% Listboxes



listSelectSheets = uicontrol('parent',figmain,'Style','Listbox','String',...
                             Sheets(1,1:end),'Units','Normalized',...
                             'Position',[0.6,0.2,0.35,0.7],'Callback',...
                             @fonclistSelectSheets);
                         
% for it to allow multiselection  

set(listSelectSheets,'Min',0,'Max',2);
                    
                

                         
listSelectFile = uicontrol('parent',figmain,'Style','Listbox','String',...
                             FILENAME,'Units','Normalized',...
                             'Position',[0.1,0.6,0.35,0.3],'Callback',...
                             @fonclistSelectFile);
                         
%% Edit Text Fields

EditNameString = 'J186G';

EditName = uicontrol('parent',figmain,'Style','edit','Units',...
                     'Normalized','Position',[0.1 0.4 0.3 0.1],'Callback',...
                     @foncEditName,'String',EditNameString);
                 
EditLocationString = 'C:\Users\Q\Desktop\Files\Internship - INL\';
                
EditLocation = uicontrol('parent',figmain,'Style','edit','Units',...
                         'Normalized','Position',[0.1 0.2 0.3 0.1],'Callback',...
                         @foncEditLocation,'String',EditLocationString);
                     
%% Text Fields

namelistSS = uicontrol('parent',figmain,'Style','text','String',...
                             'Choose which Sheets will be taken into account',...
                             'Units','Normalized','Position',...
                             [0.6,0.92,0.35,0.07],'BackgroundColor','black',...
                             'ForegroundColor','white');
                         
                         
namelistSF = uicontrol('parent',figmain,'Style','text','String',...
                     'Choose which File from which you will be selecting the sheets',...
                     'Units','Normalized','Position',...
                     [0.1,0.92,0.35,0.07],'BackgroundColor','black',...
                     'ForegroundColor','white');
                 
nameeditname =  uicontrol('parent',figmain,'Style','text','String',...
                         'Write the name of the file',...
                         'Units','Normalized','Position',...
                         [0.02,0.50,0.35,0.07],'BackgroundColor','black',...
                         'ForegroundColor','white');
                     
nameditlocation = uicontrol('parent',figmain,'Style','text','String',...
                             'Write the name of the entire Location',...
                             'Units','Normalized','Position',...
                             [0.068,0.3,0.35,0.07],'BackgroundColor','black',...
                             'ForegroundColor','white');

% This will show the selected sheets, it is Edit so the text is in the
% center, the user cannot edit since the 'Enable' property is 'inactive',
% thus the user cannot change it and it is simpl a text 

selectedSheetsListText = uicontrol('parent',figmain,'Style','edit','String',...
                                   'Nothing yet','Units','Normalized',...
                                   'Position',[0.1 0.02 0.49 0.16],...
                                   'Enable','inactive');
     

choseSheets = uicontrol('parent',figmain,'Style','text','String',...
    'The chose files are :','Units','Normalized','Position',...
    [0.01,0.10,0.08,0.1],'BackgroundColor','black',...
    'ForegroundColor','white');



                               
%% Callback functions

    %----------------------------------------------------------------------
    function foncpushNext(~,~)
        
        name = get(EditName,'String');
        location = get(EditLocation,'String');
        
        
        if (isempty(name) ==0 && isempty(location) == 0 &&...
            isempty(posSheet) == 0)
            
            % just converting into a message so the user can at
            % least understand something
            % It will be show as e.g. 1001;1001;1100 where 1 means
            % that the sheet will be taken into account, 0 it wont
            % After ; means passing to the next file
            for tt = 1:dimSheets(1)
                
                if tt == 1
                    
                    selectedSheets = num2str(posSheet(tt,1:end));
                    
                else
                    
                    selectedSheets = [selectedSheets ';' num2str(posSheet(tt,1:end))];
                    
                end
                
            end
            
            % passing from double to logical
            posSheet = logical(posSheet);
            
            % Confirm that the user indeed wants to used those
            % sheets
            
            dlgTitle = 'User Question';
            dlgQuestion = ['The selected sheets are ',...
                selectedSheets,' Do you wish to continue?'];
            choice = questdlg(dlgQuestion,dlgTitle,'Yes','No','Yes');
            
            if strcmp(choice,'Yes')
                
                uiresume(figmain);
                close(figmain);
            end
            
        elseif isempty(posSheet)
            warndlg('No Sheet has been selected')
            
        elseif isempty(location)
            warndlg('No Location has been added')
            
        elseif isempty(name)
            warndlg('No name has been added')
        end
  
    end

    %----------------------------------------------------------------------
    function foncpushCancel(~,~)
        
        Close = 1;
        uiresume(figmain);
        close(figmain);
    end


    %---------------------------------------------------------------------
    % This function saves the selected Sheets
    function foncpushSelect(~,~)
        
        indexSelectedFile = get(listSelectFile,'Value');
        indexSelectedSheet = get(listSelectSheets,'Value');
        
        % In case the user changes its mind and wants to change its sheets
        % selection, we reset the values to 0.
        
        posSheet(indexSelectedFile,1:end) = 0;
        posSheet(indexSelectedFile,indexSelectedSheet) = 1;
        
        % Write the selected positions in the text field 
        set(selectedSheetsListText,'String',mat2str(posSheet));
        
    end

    %---------------------------------------------------------------------
    % This function removes the selected Sheets
    
    function foncpushUnselect(~,~)
        
        indexSelectedFile = get(listSelectFile,'Value');
        
        % In case the user changes its mind and wants to change its sheets
        % selection, we reset the values to 0.
        
        posSheet(indexSelectedFile,1:end) = 0;
        warndlg('You just Unselected all the Sheets');
        
        % Write the selected positions in the text field 
        set(selectedSheetsListText,'String',mat2str(posSheet));
        
    end

    %----------------------------------------------------------------------
    % This function will give us the seleceted file and recover the Sheet
    % name that are inside this file. It will then make the listSelectSheet
    % data change in function of the selected file
    
    function fonclistSelectFile(hObject,~)
        
        % If it is a doubleclick
        
        if strcmp(get(figmain,'SelectionType'),'open');
            
            % To get the selected Item ; since the input cell is the
            % FILENAME, due to how the excel values were recoverd, one can
            % silply use the index to get the name of the FILE and the name
            % of the Sheets.
            
            indexSelectedFile = get(hObject,'Value');

            newSheetsName = Sheets(indexSelectedFile,1:end);
            set(listSelectSheets,'String',newSheetsName);
            
        end
        
    end

    %----------------------------------------------------------------------
    % This function will allow the user to select the sheets that he wants
    % to read and it creates a "shadow" matrix, i.e.  if it is not selected
    % at the same position in the shadow matrix it will be 0 (false), 
    % othewise it will be  1 (true). OBS : posSheets has already been
    % initialized as 0, so one only needs to change the true values
    
    function fonclistSelectSheets(hObject,~)
        

        indexSelectedFile = get(listSelectFile,'Value');
        indexSelectedSheet = get(hObject,'Value');;
        
        % In case the user changes its mind and wants to change its sheets
        % selection, we reset the values to 0.
        
        posSheet(indexSelectedFile,1:end) = 0;
        posSheet(indexSelectedFile,indexSelectedSheet) = 1;
        
        % Write the selected positions in the text field 
        set(selectedSheetsListText,'String',mat2str(posSheet));
        
    end

    %----------------------------------------------------------------------
    % This function will simply save the name chosen by the user
    
    function foncEditName(hObject,~)
    
        name = get(hObject,'Value');
        
    
    end
    %----------------------------------------------------------------------
    % This function will simply save the location chosen by the user
    
    function foncEditLocation(hObject,~)
    
        location = get(hObject,'Value');
    
    end

uiwait(figmain);

end

