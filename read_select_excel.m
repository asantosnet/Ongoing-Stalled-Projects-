function [ Data,textData,rawData,FILENAME,PATHNAME,Close,name,location,newSheetsName]...
                                                    = read_select_excel
%read_select_excel() will allow the user to select the Excel file(s) to read
%the data that will be used later on. Select only files in the same folder
%or change the code
%
%   Data will be a cell array where the data of one sheet of the excel file
%   will be in one coloom of one line and as one goes down to other lines,
%   the concerned excel file will change.
%   textData same thing as Data but for text instead of numbers, rawData
%   same thing as Data but for everything
%   FILENAME, a one line cell with the names of the excel files
%   PATHNAME, the equivalent of FILENAME but for the pathname and it is a
%   line
%   Close = 1 or 0; 1 will close the program
%   SheetsName gives back the name of the sheet, saved in the same way as
%   the DATA
%   name of the generated excel file (if needed)
%   name of the location where the generated excel file will be (if needed)

Close =0;

% The user can select multuple files

[FILENAME, PATHNAME, ~] = uigetfile('.xlsx', 'Sélect the Excel file(s)','MultiSelect','on');




% number of files selected by the user

if ischar(FILENAME)
    
    nSelectedfiles = 1;
else
    
    % In order to agree to the Data and Sheetsname cell
    
    nSelectedfiles = size(FILENAME,2);
    
end


allStatus = ones(nSelectedfiles,1);

realFile = 0;
nDeleteFiles = 0;

for i = 1:nSelectedfiles
    
    % So we don't get index out of bounds exception
    
    realFile = i - nDeleteFiles;
    
    % it can't be empty
    
    if isempty(FILENAME) == 0
        
        % Assemble the pathfile

        if nSelectedfiles == 1
            
            pathfile = fullfile(PATHNAME,FILENAME);
            
        else
            
            pathfile = fullfile(PATHNAME,FILENAME{1,realFile});
            
        end
        
        % Detect the number and names of the sheets and if it is a escel
        % spreadsheet
        
        [status,sheets] = xlsfinfo(pathfile);
        
        % Save if it is readable or not
        
        allStatus(realFile,1) = ~strcmp(status,'Microsoft Excel Spreadsheet');
        
        % Saving the names of the sheets
        
        % Checking for erros - it will be there if sheets is a char 
        
        if ischar(sheets) == 0
            
            if realFile == 1
                
                SheetsName = sheets;
                
            else
                
                % to avoid dimension problems
                
                ncolSheetsName = size(SheetsName,2);
                ncolsheets = size(sheets,2);
                
                if ncolSheetsName > ncolsheets
                    
                    % to get it to the proper dimension
                    
                    sheets = [sheets cell(1,(ncolSheetsName - ncolsheets))];
                    
                elseif ncolSheetsName < ncolsheets
                    
                    nlinSheetsName = size(SheetsName,1);
                    SheetsName = [SheetsName cell(nlinSheetsName,...
                        (ncolsheets-ncolSheetsName))];
                    
                end
                
                SheetsName = [SheetsName;sheets];
                
                
            end
        else
            
            warndlg(sheets);
            warndlg('This file will be removed');
            
            
            % This works only if FILENAME is a line cell,e.g. cell(1,8)
            FILENAME(realFile) = [];
            nSelectedfiles = nSelectedfiles -1;
            nDeleteFiles = nDeleteFiles +1;
            
            
            
        end
    else
        
        warndlg('You haven''t sellected anything')
        
    end
    
end

% The user will select what he wants or not to be read, the name and
% location of the possibly generated excel file

[name,location,posSheet,Close ] = selectExcelSheet(SheetsName,FILENAME);

nNonSelectedFiles = 0;
realrow = 0;

Data = cell(nSelectedfiles,1);
textData = cell(nSelectedfiles,1);
rawData = cell(nSelectedfiles,1);
if Close == 0

    % now one will read the files

    for g = 1:nSelectedfiles

        % Verify if the file is indeed readable

        if allStatus(g,1) ==0
            
            realrow = g - nNonSelectedFiles;

            tempSheets = SheetsName(realrow,1:end);

            % removing the unwanted sheets, in this case it needs to be g
            % and not realrow

            tempSheets = tempSheets(posSheet(g,1:end));
      
            nSheets = size(tempSheets,2);
           
            % Assemble the pathfile
            
            if nSelectedfiles == 1
                
                pathfile = fullfile(PATHNAME,FILENAME);
                
            else
                
                pathfile = fullfile(PATHNAME,FILENAME{1,realrow});
                
            end            
            
            % In case no Sheets have been selected
            if nSheets ~= 0
                
                % Recover the data from each sheet
                for t = 1:nSheets;

                    % Recover the data from the sheets and also the used
                    % sheets name
                    
                    [Data{realrow,t},textData{realrow,t},rawData{realrow,t}]...
                        = xlsread (pathfile,tempSheets{t});
                    newSheetsName{realrow,t} = tempSheets{t};
                end
            else
                
                % If no Sheets have been selected, remove FILENAME,
                % and the Data position corresponding to that
                % data ( To avoid any future bugs)
                
                FILENAME(realrow) = [];
                
                tempData = Data(1:realrow,1:end);
                tempData(all(cellfun('isempty',tempData),2),:) = [];
                
                if g == nSelectedfiles
                    
                    Data = tempData;
                    
                elseif g==1
                    Data = [tempData;Data((realrow+1):end,1:end)];

                else
                    Data = [tempData;Data((realrow+1):end,1:end)];

                end
                
                nNonSelectedFiles = nNonSelectedFiles +1;
                
            end
        else

            warndlg('Can''t be read by xlsread')

        end
    end

end

end

