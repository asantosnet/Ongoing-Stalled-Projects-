function [excel,eWorkSheets,eWorkbook] = plotXYexcel(ColumnX,ColumnY,X,Y,Data, FILENAME,...
                                                     location,CreatVect,excel)
%plotIVexcel(Data, FILENAME ) This function plots I in function of V for a
%given Data set and FILENAME 
%   Data, the values recovered with read_select_excel
%   FILENAME - Idem
%   ColumnY and V where in the Data or in the excel file the I and V data
%   is, in which column they are ( in the excel file for example)
%   name, the name of the excel file that will be generated
%   this creat one file with several worksheets, each saving some results
%   that correspond to the plot later on
%   The data to be used in case a I_V - the Columns 1 which is Id and 5 is V
%   manage the data so one can read only the I and V columns
%   DO NOT REMOVE, always after a write_excel, when stoped using the workbook,
%   write the folowing code. This function uses write_excel
%       
%       eWorkbook.Close;
%       excel.Quit;
%       delete(excel);
%
%   e.g. ColumnY = 1;ColumnX = 5;name = 'J186H' ;
%        location =  'C:\Users\Q\Desktop\Files\Internship -
%        INL\Codigo_Matlab\';
%        X and Y the names of the X and Y datas , e.g. X = Vd(V); Y = Id(A)



nExcelFiles = size(Data,1);

[ DataXY,HeaderXY,RangeXY,FilesXY]...
         = plotXYData(Data,FILENAME,ColumnY,ColumnX,X,Y);

% Since write_excel only accepts cells

nameSheet = {['Plot - ' Y '_' X]};
     

% For each excel file, I will create a new excel in which the plot will be
% located. For each type of plot, in this case XY, one sheet will be 
% created, where each 2 columns of the excel file ( since X and Y) 
% corresponds to the data from one sheet. All the important data from the
% excel file will be in this one sheet, the plot will be made in a new
% sheet.
% Orginal name of excel file + 'result'
% If the excel file arealdy exists , then a new sheet will be
% created with the name of the manipulation that was made to it.


for i = 1:nExcelFiles
    
    % If i>1 then we will already have created an actxserver, there is no
    % need to intialize a new one
    
    if i >1
        
        CreatVect(1,1) = 0;
        
    end
    
    name = FilesXY{1,i};
    
    % Creating/opening the Excel file where the data will be saved, the
    % nescessary Sheet will also be created/open
    dbstop in write_excel_file
    [excel,eWorkSheets,eWorkbook] = write_excel_file(RangeXY(i,:),...
        DataXY(i,:),name,nameSheet,location,CreatVect,excel,HeaderXY(i,:));

%     % plot Y vs X
% 
%     inputdlg(['Save/overide the excel file and close any excel popups'...   
%               'in case pirated version']),   
%    
%     
%     if ischar(FILENAME)
%     
%         nameWorksheet = FILENAME;
%     
%     else
%     
%         nameWorksheet = FILENAME{i};
%     end
%     
%     % This way makes it harder to manipulate the Data on excel if needed
%     % DataXYplot = DataXY(i,1:end);
%     % Thus,the harder, but better way is to recover the data from the
%     % columns that we wrote in in the last function
%      
%     sheet = eWorkSheets.Item(nameWorksheet);
%     
%     dbstop in range
%     
%     range = RangeXY(i,1:end);
%     
%     % remove any empty cells
%     range = range(~cellfun('isempty',range));
%     
%     
%     rangeSize = size(range,2);
%     DataXYplot = cell(1,rangeSize);
%     
%     for t = 1:rangeSize
%         
%         % I need to do this since range2 = {'fvfg'} and range2{1} = fvfg 
%         range2 = range{t};
%         
%         % Le B1 est une sring, alors il fauut commencer depuis le B2 pour que
%         % ça ne buge pas la plot. I know that the first value of the Column
%         % is always a string since it the excel file created by this
%         % program. Additionally, range2 = 'B1:B202', range3 = B1:B202
%         
%         range3 = range2{1};
%         range3 = [range3(1) '2' range3(3:end)];
%         
%         DataXYplot{1,t} = sheet.Range(range3);
% 
%     end
%     
%     % The name of the File ends with .xlxs or .xls, therefore I will remove
%     % the last 4 characters.
%     
%     title = strcat(X,'_',Y,'_','for','_',nameWorksheet(1:(end-4)));
%     xtitle = X;
%     ytitle = Y;
%     
%     dbstop plot_excel
%     
%     plot_excel (eWorkbook,eWorkSheets,excel,title,xtitle,ytitle,DataXYplot,nameWorksheet)
  
    
end

end

