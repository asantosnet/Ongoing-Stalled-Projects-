function [ DataXY,HeaderXY,RangeXY,name,sheetsXY,location ]...
         = plotXYData(Data,FILENAME,ColumnY ,ColumnX,name,location,Nplots,X,Y)
%plotXYData(Data,FILENAME,ColumnY ,ColumnX,name,location,Nplots,X,Y)
%This arranges the data to send to the write_excel fuunction
%   Data the one that came from the read_select_excel
%   ColumnY  and ComunX corresponds to the column in the excel file from
%   which the Data gotta be taken and FILENAME the one originated from
%   read_select_excel
%   
%   e.g Header = {'tt' 'to' 'ti' 'te';'a' 'b' 'v' 'd'};
%       Range = {'A1:A3' 'B1:B3' 'C1:C3';'A1:A3' 'B1:B3' 'C1:C3'};
%       name = 'teste';
%       sheets = {'teste1' 'tests2'};
%       location = 'C:\Users\Q\Desktop\Files\Internship - INL\Codigo_Matlab\';
%       Data = cell(2,3);
%       Data{1,1} = [1;2];
%       Data{1,2} = [1;3];
%       Data{1,3} = [3;3];
%       Data{2,1} = [1;2];
%       Data{2,2} = [1;3];
%       Data{2,3} = [3;3];
%       ColumnY  = 1;ColumnX = 5;name = 'J186H' ;
%        location =  'C:\Users\Q\Desktop\Files\Internship -
%        INL\Codigo_Matlab\';
%        Nplots = The number of pltos that were chosen during
%        choose_manip
%        X and Y the names of the X and Y datas , e.g. X = Vd(V); Y = Id(A)


% the number of files

nExcelFiles = size(Data,1);


DataXY = cell(nExcelFiles);
HeaderXY = cell(nExcelFiles);
RangeXY = cell(nExcelFiles);
sheetsXY = cell(1,nExcelFiles);



for t = 1:nExcelFiles
    
    if ischar(FILENAME)
    
        sheetsXY{1,t} = FILENAME;
    
    else
    
        sheetsXY{1,t} = FILENAME{t};

    end

    
    % Deal with void / Gotta do it here / FIND BETTER SOLUTION
    % remove void cells to avoid erros, I don't want to use for, but I gotta
    % faster to use 'isempty' than @isempty

    Datafile = Data(t,1:end);
    
    % detect empty spaces in the line
    Index = cellfun('isempty',Datafile);
    
    % remove those empty spaces
    Datafile(Index) = []; 
    
    nSheets = size(Datafile,2);
    
    % Position where the I data and V data will be saved in the excel file
    Ypos = (1:2:((2*Nplots + 2)*nSheets));
    Xpos = (0:2:((2*Nplots + 2)*nSheets));
    
    b = 1;
    
    for g = 1:nSheets
        
        % Assemble the current and voltage values
        DataSheet = Datafile{1,g};
        DataY = DataSheet(1:end,ColumnY ); % take the first column
        DataX = DataSheet(1:end,ColumnX); % take the second column
        
        DataXY{t,Ypos(g)} = DataY;
        DataXY{t,Xpos(g+1)} = DataX;
        
        % Define the Headers        
        headerY = strcat(Y,'_',int2str(g));
        headerX = strcat(X,'_',int2str(g));
        
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
        % Data for the second plot, then it used 2 columns, then jump another
        % two columns to write the Data from the second sheet if Nplots = 2
        
        b = b + 2 + (2*Nplots -2);
        
        
    end

end



end

