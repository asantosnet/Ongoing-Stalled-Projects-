function [eSheet,eWorksheets] = create_Sheet( sheetName,excelObject)
%[eSheet] = creat_Sheet( sheetName,exelObject)
%   DThis function creats a new Excel sheet with the name sheetName
%   sheetName is a string and the name of the Sheet
%   excelObject can be a Workbook or a Worksheets. In the former a new sheet
%   will be added an a Worksheet initialized. For the latter case a
%   worksheet will be added after the last worksheet in the Worksheets
%   object
%   eSheet is the created sheet
%   eWorksheet is the used Worksheets


% Creating the new Sheet

if strcmp(class(excelObject),'Interface.Microsoft_Excel_15.0_Object_Library._Workbook')
    
    eWorksheets = excelObject.Worksheets;
    eSheet = eWorksheets.Item(1);
    
elseif strcmp(class(excelObject),'Interface.Microsoft_Excel_15.0_Object_Library.Sheets')
    
    eWorksheets = excelObject;
    eSheet =  eWorksheets.Add([],eWorksheets.Item(eSheets.Count));
    
else
    
    error(['Error. \nexcelObject is : ' class(excelObject) ' - This is not permitted']);
    
end


% Certain names are not allowed
try
    
    eSheet.Name = sheetName;
    
catch errorName
    
    if strcmp(errorName.identifier,'MATLAB:COM:E2148140012')
        % Remove any char that can't be used as a name
        % and reduce its size to
        % 31 char 
        
        sheetName = regexprep(sheetName,'[:\/?*[]]','');
        
        if size(sheetName,2) > 31
            sheetName = sheetName(1,1:31);
        end
        eSheet.Name = sheetName;
    else
        rethrow(errorName);
    end
    
end

end

