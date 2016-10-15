function [ eSheet,ignoreSheet] = find_sheet(sheetName,eWorkSheets)
%[ eSheet,ignoreSheet] = find_sheet(sheetName,eWorkSheets)
%   This function will recover the Sheet from a give WorkSheet.
%   sheetName is the name of the wanted sheet (string)
%   eWorkSheets is the e.Workbook.Worksheets object where the sheet is
%   located. In case no sheet is located, there is a option of creating a
%   new one.
%   Returns the object sheet and if there was an error in the process, if
%   this particular sheet should be ignored or not.


ignoreSheet =0;

% The file migth not exist for example
try
    
    eSheet = eSheets.Item(sheetName);
    
catch errorItem
    
    % If the error is the one expected
    if strcmp(errorItem.identifier,'MATLAB:COM:E2147614731')
        
        % It will creat another sheet if the user so desires
        
        dlgTitle = 'User Question';
        dlgQuestion = ['Do you want to create a new sheet?'...
            'Failing to do so will cause this sheet to be igonred'];
        choice = questdlg(dlgQuestion,dlgTitle,'Yes','No','Yes');
        
        if strcmp(choice,'Yes')
            
            % Creating the new sheet
            
            [eSheet,~] = create_Sheet( sheetName,eWorkSheets);
          
            
        elseif strcmp(choice,'No')
            
            
            % Tell what is wrong and move to the other iteration of the
            % for loop
            warndlg([errorItem.identifier;erroItem.message;...
                erroItem.cause]);
            wanrdlg('This sheet will be ignored');
            
            ignoreSheet =1;
            return; % leave function
            
        end
        
    else
        
        % Tell what is wrong and move to the other iteration of the
        % for loop
        
        warndlg([errorItem.identifier;erroItem.message;...
            erroItem.cause]);
        wanrdlg('This sheet will be ignored');
        ignoreSheet =1;
        return; % leave function
    end
end

end

