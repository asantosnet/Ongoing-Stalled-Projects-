
function [eWorkbook,changeCreatVect] = find_workbook (e,name,location)
%[eWorkbook,changeCreatVect] = findtheworkbook (e,name,location)
% Verify if the ActiveWorkbook is the one that we wanna work with,
% otherwise, open a new Woorkbook, Otherwise, creat a new one.
% eWorkbook is the ( already active) workebook that can be worked with( the
% one that is named name)
% location is the Path to where the excel file is saved
% changeCreactVect in case it is being used by write_excel, read more
% information there


eWorkbook = e.ActiveWorkbook;
eWorkbooks = e.Workbooks;


if strcmp(name,eWorkbook.Name) == 0
    
    % Verify if a workbook with this name is already open, if not,
    % we will need to open it or creat another one
    try
        eWorkbook = eWorkbooks.Item(name);
        
    catch errorItem
        
        % If the error is the one expected
        if strcmp(errorItem.identifier,'MATLAB:COM:E2147614731')
            
            % It will try to open the wanted excel file
            
            [eWorkbook,changeCreatVect] = open_workbook(e,location,name);
            
        else
            rethrow(errorItem)
        end
    end
    
    eWorkbook.Activate;
    
end

end

