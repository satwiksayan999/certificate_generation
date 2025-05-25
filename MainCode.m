clc 
clear all
close all 
home 

filename = 'Registration_Details.xls';
[num,txt] = xlsread(filename)



len=length(txt)


blankimage = imread('Certificate_Blank.tif');


for i=1:len
    for j= 2:2 
        text_names(i,j)=txt(i,j)
    end
end


for i=1:len
    for j= 3:3
      text_topic(i,j)=txt(i,j)
    end
end




for i=2:len
        text_all=[text_names(i,2) text_topic(i,3)]
       
        
        position = [100 258;120 416];         
        
        
        RGB = insertText(blankimage,position,text_all,'FontSize',22,'BoxOpacity',0);
        
        
        figure
        imshow(RGB)        
        
        
        y=i-1
        filename=['Certificate_Topic_' num2str(y)]
        lastf=[filename '.tif']
        saveas(gcf,lastf)
        % generate and save files with .tif extension
end    

% Code works for any length of data ROWWISE AND COLUMNWISE
% Last for loop with minor changes facilitate new project implemntations


