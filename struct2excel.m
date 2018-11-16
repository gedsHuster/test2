%% 将结构体保存到excel
clc;clear; 
s.China = [1 2 3]'; 
s.US = [1 2 3 4]'; 
s.UK = [1 2 3 4 5]';
cell = struct2cell(s);% struct转成cell
for sheet=1:3
    xlswrite('output_struct.xlsx', cell{sheet}, sheet);
end
sheet_name = fieldnames(s);
%% 改sheet名字
e = actxserver('Excel.Application'); % # open Activex server
ewb = e.Workbooks.Open('C:\Users\geds\Desktop\output_struct.xlsx'); % # open file (enter full path!)
for sheet = 1:3
    ewb.Worksheets.Item(sheet).Name = sheet_name{sheet}; % # rename 1st sheet
end
ewb.Save % # save to the same file
ewb.Close(false)
e.Quit