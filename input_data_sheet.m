function input_data_sheet(eSheet,Range,Data,varargin)
%input_data_sheet(eSheet,Range,Data)
%input_data_sheet(eSheet,Range,Data,Header)
%   This function writes the data in a Excel sheet
%
%   INPUT
%
%   eSheet is the excel " object" that contains the sheet were the data
%   will be written.
%
%   Header contains the String that will be written in the first row of
%   each column. Header is a line cell where each column is related to one
%   excel column
%
%   Range contains the row and column where each data will be writen. Range
%   is a is a line cell where each column is related to one excel column.
%
%   Data contains the Data itself.Data is a line cell where each column is 
%   related to one excel column

%% Checking inputs*

switch nargin
    
    case  3
        
        Header ='';     
        
    case 4
        
        if iscell(varargin{1})
            
            Header = varargin{1};
            
        else
            error('Error. \nClass of last input is not accepted')
        end
        
    otherwise
        error('Error. \nToo many/few input arguments')
end


%% Begining of the code

% Activate the Sheet
eSheet.Activate

% Remove empty spaces
IndexRange = cellfun('isempty',Range);

% remove those empty spaces
Range(IndexRange) = [];

nCollomns = size(Range,2);

% Writing the Data into each collomn

for g = 1:nCollomns
    
    inputData = Data{1,g};
    inputData = num2cell(inputData);
    
    % input the data
    
    Range2 = Range{g};
    
    eActivesheetRange = get(eSheet,'Range',Range2{1});
    
    if iscell(Header)    
        eActivesheetRange.Value = [Header{g};inputData];
    else
        eActivesheetRange.Value = inputData;
    end
    
end

end

