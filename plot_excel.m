function  plot_excel (eWorkbook,eWorkSheets,excel,title,xtitle,ytitle,Data,nameWorksheet )
%plot_excel (eWorkSheet,title,xtile,ytitle,Data ) Plots something on an
% excel sheet after the chosen sheet
%   eWorksheet the worksheet object from the opend workbook
%   title the title of the plot and the sheet name
%   xtitle and ytitle the, respectevely, x and y axis title
%   Data, a cell containing the data to be plotted in each colomn. Colmon
%   1 plotted in function of 2, 3 in function of 4 etc.. All the data is
%   plotted in one graph. i .e. 1 the y and 2 the x
%   sheetnumber, the number after which the sheet with the graph will be
%   added                              


ncolomns = size(Data,2);

% Recover the index of the worksheet that contains the Data that we want to
% plot

posSheet = eWorkSheets.Item(nameWorksheet).Index;

% verify if we have the proper data set, i.e. one Column for I and another
% for V, pair number of columns

if mod(ncolomns,2) == 0
    
    warndlg('bad Data set');
    
    % Leaves the function if the Dataset is bad
    return;
    
end


% Add a Sheet where the plot will be added


eWorkSheets.Add([],eWorkSheets.Item(posSheet),1);

% recover the created sheet and activate it

eSheets = excel.ActiveWorkbook.Sheets;

eSheetplot = eSheets.get('Item',(posSheet+1));


eSheetplot.Activate

% rename it
% Remove any char that can't be used as a name and reduce its size to
% 31 char

titleAdapted = regexprep(title,'[\/?*[]]','');

if size(titleAdapted,2) > 31
    titleAdapted =titleAdapted(1,1:31);
end

% This name migth have already been used due to the size reduction,
% thus we will use try and catch to assign a new name so the user knows
% what the plot means

try
    
    eSheetplot.Name = titleAdapted;
    
catch
    
    warndlg(['Problem giving a title to the file where the plot will be'...
        ', it migth already exist :' titleAdapted]);
    
    % Ask for the user to rename the plot manually
    answer = inputdlg([' The original name was : ' title ' - Please rename'...
        ' it as you wish, chose it correctly or leave it empty']);
    
    if isempty(answer{1}) == 0
        eSheetplot.Name = answer{1};
    end
end



% Add a chart

eChartObjetc = eSheetplot.ChartObjects.Add(400,30,500,500);
chart = eChartObjetc.Chart;

%chart.Name = title;

%chart.Activate;




% Position where the x and y data will be saved in the excel file

ypos = (1:2:(3*ncolomns));
xpos = (0:2:(3*ncolomns));

% Plot the chosen Colmuns
for i = 1:(ncolomns/2)
    
    % Creating a new plot inside the same chart
    SeriesName = strcat('mesurement',int2str(i));
    
    chart.SeriesCollection.NewSeries;
    series = chart.SeriesCollection(i);
    
    % Selecting the data range
    
    series.Value = Data{1,ypos(i)};
    series.XValue = Data{1,xpos(i+1)};
    
    series.Name = SeriesName;
    
end


% Set Chart type

chart.ChartType = 'xlXYScatterLinesNoMarker';

% Set the X and Y titles
% axes(1) primary x, axes(2) primary 2, if secondat axes(2,1) and
% axeis(2,2) gives the x,y secondary



chart.Axes(1).HasTitle = 1;
chart.Axes(1).AxisTitle.Text = xtitle;

% Set the Y values


chart.Axes(2).HasTitle = 1;
chart.Axes(2).AxisTitle.Text = ytitle;

% Title of the chart

chart.HasTitle = 1;
chart.ChartTitle.Text = titleAdapted;


% Save what migth be important information on the first column of the
% first row
% input the data

eActivesheetRange = get(eSheetplot,'Range','A1');

eActivesheetRange.Value = title;

%Save

invoke(eWorkbook,'Save');

end

