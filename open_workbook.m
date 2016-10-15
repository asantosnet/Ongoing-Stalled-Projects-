function [eWorkbook,changeCreatVect] = open_workbook(e,location,name)
%[eWorkbook,changeCreatVect] = open_Workbook(location,name)
%   This fonction attempts to open a excel file using eWorkbooks, it deals
%   with the case that it is impossible to open it and gives the user the
%   option of creating a new excel file

changeCreatVect = 0;
eWorkbooks = e.Workbooks;

try
    eWorkbook = eWorkbooks.Open(fullfile(location,name));
    
    % If the workbook cannot be opened
catch errorOpen
    
    warndlg([errorOpen.identifier;errorOpen.message...
        errorOpen.cause]);
    
    dlgTitle = 'User Question';
    dlgQuestion = ['Do you want to create a new excel file ?'...
        'Failing to do so will crash the program.'];
    choice = questdlg(dlgQuestion,dlgTitle,'Yes','No','Yes');
    
    if strcmp(choice,'Yes')
        
        % Creat the workbook
        eWorkbook = eWorkbooks.Add;
          
        %reset CreatVect(1,2) to 1
        changeCreatVect = 1;
        
        
    elseif strcmp(choice,'No')
        
        rethrow(errorOpen)
        
    end
end

eWorkbook.Activate;

end


