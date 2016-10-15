function [ DataXY,HeaderXY,RangeXY,FilesXY]...
         = plotXYData(Data,FILENAME,ColumnY ,ColumnX,X,Y)
%plotXYData(Data,FILENAME,ColumnY ,ColumnX,name,location,X,Y)
%This arranges the data to send to the write_excel function
%   Data the one that came from the read_select_excel
%   ColumnY  and ComunX corresponds to the column in the excel file from
%   which the Data gotta be taken and FILENAME the one originated from
%   read_select_excel. 
%   DataXY is a cell, each row corresponds to one file, each column of that
%   row corresponds to one column vector that will be writen in the excel
%   file. In this way, the position where the vector will be written is
%   saved in the same cell index as where the corresponding vector is
%   saved.
%   HeaderXY is the header of each X and Y column written in the first row
%   of the excel file, arranged in the same way as for the Data
%   RangeXY is the range corresponding to the rows where the X and Y data 
%   will be written in the excel file
%   FilesXY is a cell(1,n) that contains the name of the files that will be
%   created. The name of the sheets is not written here since they are the
%   same for all files, which is, the name of the function + the X and Y
%   strings
%   
%   e.g Header = {'tt' 'to' 'ti' 'te';'a' 'b' 'v' 'd'};
%       Range = {'A1:A3' 'B1:B3' 'C1:C3';'A1:A3' 'B1:B3' 'C1:C3'};
%       location = 'C:\Users\Q\Desktop\Files\Internship - INL\Codigo_Matlab\';
%       Data = cell(2,3);
%       Data{1,1} = [1;2];
%       Data{1,2} = [1;3];
%       Data{1,3} = [3;3];
%       Data{2,1} = [1;2];
%       Data{2,2} = [1;3];
%       Data{2,3} = [3;3];
%       ColumnY  = 1;ColumnX = 5;
%        location =  'C:\Users\Q\Desktop\Files\Internship -
%        INL\Codigo_Matlab\';
%        X and Y the names of the X and Y datas , e.g. X = Vd(V); Y = Id(A)


% the number of files

nExcelFiles = size(Data,1);


DataXY = cell(nExcelFiles);
HeaderXY = cell(nExcelFiles);
RangeXY = cell(nExcelFiles);
FilesXY = cell(1,nExcelFiles);


for t = 1:nExcelFiles
    
    % Datafile is equal to the Data row corresponding to the desired FIle
    
    Datafile = Data(t,1:end);
    
    % Verify if the Data row is not entirely empty, if it is, the function
    % will continue to its next iteration
    
    if any(~cellfun('isempty',Datafile)) == 0
        
        % To pass to the next value of t
        continue;
        
    end
    
    
    if ischar(FILENAME)
        
        % The .xlsx need to be maintained at the end
        FilesXY{1,t} = ['results_' FILENAME];
    
    elseif iscell(FILENAME)
    
        FilesXY{1,t} = ['results_' FILENAME{t} ];

    end
   
    % Detect empty spaces in the line and remove them
    
    Datafile(cellfun('isempty',Datafile)) = []; 
    
    % The number of sheets in the excel file, resulting in nSheetsx2 used
    % columns
    
    nSheets = size(Datafile,2);
    
    % Position where the X data and Y data will be saved in the excel file
    % E.g. firs row X_1 second row Y_1, 5 row X_2, 6 row Y_2. 
    
    Ypos = (1:2:(4*nSheets));
    Xpos = (0:2:(4*nSheets));
    
    b = 1;
    
    for g = 1:nSheets
        
        % Assemble the X and Y values
        DataSheet = Datafile{1,g};
        DataY = DataSheet(1:end,ColumnY ); % take the first column
        DataX = DataSheet(1:end,ColumnX); % take the second column
        
        DataXY{t,Ypos(g)} = DataY;
        DataXY{t,Xpos(g+1)} = DataX;
        
        % Define the Headers        
        headerY = strcat(Y,' - Sheet ',int2str(g));
        headerX = strcat(X,' - Sheet ',int2str(g));
        
        HeaderXY{t,Ypos(g)} = headerY;
        HeaderXY{t,Xpos(g+1)} = headerX;
        
        % XLSCOLNUM2STR(1) gives 'A' vice versa gives 1
        % Define the range of the column, i.e. how many lines will be used
        ndY = size(DataY,1);
        ndX = size(DataX,1);
        
        startY = strcat(xlsColNum2Str(Ypos(b)),'1');
        startX = strcat(xlsColNum2Str(Xpos(b+1)),'1');
        
        endY = strcat(xlsColNum2Str(Ypos(b)),int2str(ndY));
        endX = strcat(xlsColNum2Str(Xpos(b+1)),int2str(ndX));
        
        rangeY = strcat(startY,':',endY);
        rangeX = strcat(startX,':',endX);
        
        RangeXY{t,Ypos(g)} = rangeY;
        RangeXY{t,Xpos(g+1)} = rangeX;
        
        % We are going to plot in e.g. A AND B, jump 2 for writting the
        % Data for the second sheet, then it used 2 columns, then jump another
        % two columns to write the Data from the third sheet if nSheets = 3
        
        b = b + 2 + (2*nSheets -2);
        
        
    end

end



end

