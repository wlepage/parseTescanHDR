%% config
% data_dir = 'W:\Experiments\20160622\';  % folder that has the Tescan header files
data_dir = 'C:\Users\wlepage\Google Drive\Dasgupta-Battery Subgroup\Data_Characterization\SEM\20170201 DIC patterned Li\'; 
save_filename = 'headers_20170201'; % name of the Excel file output

%% prep and preallocate
main_dir = pwd;
cd(data_dir); % go to the folder with the hdr files
filestruct = dir('*.hdr'); % make a file structure of the hdr files
filecount = length(filestruct);
outputs = cell(filecount+1,63); % there are 62 lines of data in each header file, if line/frame accumulation and dynamic focus are off
outputs{1,1} = 'filename';

%% loop through files
for ii = 1:filecount
    % load the text
    fid = fopen(filestruct(ii,1).name);
    C = textscan(fid, '%s%s', 63, 'HeaderLines',1,'Delimiter','=','CollectOutput',false);
    fclose(fid);
    if ii==1 % for the first file only, make the header row of the spreadsheet
        if strcmp(C{1,1}{2},'Date') % if there was not line or frame accumlation
            % then add the line/frame accumulation type entry
            C{1,1} = [C{1,1}{1} {'AccType'} C{1,1}{2:end}];
        end
        if not(strcmp(C{1,1}{25},'DynamicFocusAngle1')) % if there was not dynamic focus
            % then add the dynamic focus entry
            C{1,1} = [C{1,1}{1:24} {'DynamicFocusAngle1','DynamicFocusAngle2','DynamicFocusBreakPos'} C{1,1}{25:end}];
        end
        for ij = 1:length(C{1,1})
            outputs{1,ij+1} = C{1,1}{ij};
        end
    end
    outputs{ii+1,1} = filestruct(ii,1).name; % add the filename to the first column
    % condition the second row from the text files
    if strcmp(C{1,2}{1},'1') % check if the image did not use line/frame accumulation
        C2{1,2} = C{1,2};
        C2{1,2}{2} = 'none';
        for jj = 2:60 % shift the lines down
            C2{1,2}{jj+1} = C{1,2}{jj};
        end
        C{1,2} = C2{1,2};
    end
    if strcmp(C{1,2}{26},'Schottky') % check if the image did not use dynamic focus
        C2{1,2} = C{1,2};
        for jj = 25:61
            C2{1,2}{jj+3} = C{1,2}{jj};
        end
        C2{1,2}{25} = {''};
        C2{1,2}{26} = {''};
        C2{1,2}{27} = {''};
        C{1,2} = C2{1,2};
    end  
    for ik = 1:length(C{1,2})
        outputs{ii+1,ik+1} = C{1,2}{ik}; % format the cell to be writted to the spreadsheet
    end
end
xlswrite([save_filename,'.xlsx'],outputs);
cd(main_dir); % go back to the starting directory