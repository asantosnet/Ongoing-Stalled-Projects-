% This program will call allow the user to handle and add functions that
% will use the DATA obtained from the chosen excel files/sheets

close all
clear all
Close = 0;


%% To avoid bugs

dlgTitle = 'User Question';
dlgQuestion = ['Have you already closed all Exel files that you are going  ',...
               ' to modify/work on?'];
choice = questdlg(dlgQuestion,dlgTitle,'Yes','No','Yes');

if strcmp(choice,'No')
    warndlg('The application will close');
    Close = 1;
else
    warndlg('Do not close the Excel files that migth open');
end



%% recovering the data from the chosen files

if Close == 0

[Data,textData,rawData,FILENAME,PATHNAME,Close,name,location,SheetsName] =...
                                                        read_select_excel;


end

%% Choose which manipulations ( with the data) will the user be able to choose from
%  and evalute his choice. All functions that have the same type will have
%  the same input and output. Next to plot, the number of inputs and
%  outputs is written e.g. [1 2]

Availablefunc = {'plotXYexcel' 'Plot' [9 3]};
AvailablefuncPanel ={'panel_plotXY' 'Plot' [4 5]};

if Close == 0
   
[Chosenfunc,Inputfunc,Close] = choose_manip(Availablefunc,...
                                AvailablefuncPanel,Data,SheetsName,FILENAME);

end

% This realtes to how the write_excel funtion will behave, more information
% see write_excel. If CreatNewExcel is 1 then the rest needs to be one,
% same thing for Workbook and Worksheet

CreatNewExcel = 1;
CreatNewWorkSheet = 1;
CreatNewWorkbook = 1;

CreatVect = [CreatNewExcel CreatNewWorkbook CreatNewWorkSheet];

nDataManip = size(Chosenfunc,1);

if Close ==0
    
    for t = 1:nDataManip
        
    % Verify the type of manip
    % look for the index on the Availablefunc and then use this
    % index to recover the name of the function and its type
    
    NameFunc = Chosenfunc{t,1};        
    indexSelectedfunc = find(ismember(Availablefunc(1,1:2),NameFunc));
    FuncType = Availablefunc{indexSelectedfunc,2};
    NInputOutput = Availablefunc{indexSelectedfunc,3};
    
    switch FuncType
        case 'Plot'
            
            %% Plot and save the data used in a PLOT type function 
            
            % Number of inputs and outputs required for a Plot type function        
            
            Input = cell(1,NInputOutput(1));
            
            % There are 5 inputs that need to be recovered from the Input
            % cell originated from choose_manip. The rest are datas that
            % need, until further notice, be given manually
            
            for tt = 1:(NInputOutput(1)-4)
                
                if isempty(Inputfunc{t,tt}) == 0
                    
                    Input{tt} = Inputfunc{t,tt};
                    
                else
                    warndlg(['Error in : ' NameFunc ' - of type :  '...
                            FuncType ' - during ' ' input parameters']);
                end
            end
            
            % Adding the manual parameters
            
            Input{6} = FILENAME;
            Input{7} = location;
            Input{8} = CreatVect;
            
            %     ColumnY = 1;
            %     ColumnX = 5;
            %     Nplots = 1;
            %     X = Vd(V);
            %     Y = Id(A); 
            %     name = 'J186H';
            %     location = 'C:\Users\Q\Desktop\Files\Internship - INL\Codigo_Matlab\';
            
            
            % In case you want to initialize actxserver
            if CreatVect(1,1) == 1
            
                Input{9} = 'excel';  
            
            else
                
                Input{9} = excel; 
            
            end

            dbstop in plotXYexcel
            [excel,eWorkSheets,eWorkbook] = plotXYexcel(Input{1},Input{2},...
                Input{3},Input{4},Input{5},Input{6},Input{7},Input{8},...
                Input{9});
            
        otherwise
            
            warndlg('Type doesn''t exist, problem in function delcaration')
    end
    
    %DO NOT REMOVE, always after a write_excel, when stoped using the workbook,
    %wirte the folowing code

    eWorkbook.Close;
    excel.Quit;
    delete(excel);
    
    
    
    
    end
end





%% KNOWN BUGS

% When an excel file is open it migth give a Inoke Erros, Dispatch
% Exception
% When the excel file is closed, it bugs everything, I would have to reopen
% it, I don't. This can indeed be fixed by using a catch i guess
% Fix bug in Data collection for the I_V plot
% Fix bug in the panel-plot after saving and the fact that X and Y are not
% obtained
% Sometimes, the entire second row B is replace by the Title, this is
% completly random for now, find out why this happens


%% Improvements
% Speed : during read_select_excel stop using xlsread for each sheet, it
% takes too long ( 25s more or less per call)
% Change the crappy solution for the problem in panel plot 

% Show the real sheet names not only the 1 and 0
% Create more Panel types
% open the worksheet outside of the plot function if possible
% Verify that the thing to check if one needs to Creat another sheets etc... in the
% plotXYexcel file works
% Fiw how input is written so it is easy to write for all plot type
% functions
% Remove the nescessity of providing inputs manually to the function in
% main
% See if using Sheets.Add it is possible to select the type of sheet you
% are creating
% Give the guy the option to save twice 
% Make it work for the kind of excel file we have here and then
% generalize,i.e. assume that numbers, names, etc.. migth be in there too


%% DOUBTS
% The figure
% By doing this I am storing all the data in the figmain variable ( same
% role as handle maybe???, if I need to add any additional properties,
% there are ways. To find the handles e.g. findall, I don't get it very
% well yet


