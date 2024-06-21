function [] = vpextract()
% vpextract.m
% Author: Louis Leung, Ph.D.
% Department of Psychiatry & Behavioral Sciences, Stanford University
%    
% This program extracts columns from viewpoint .xlsx file.
% 1. Convert the excel file from .XLS -> .xlsx using saveas in Excel
% 2. Select parameters and plot types for your analysis. 
%   Plots are saved to folder 'Summary'.
% 3. Choose whether to create combined video output for review.
%   Videos are saved to folder 'Exported Videos'.
%
% Dependencies: This program requires the following scripts:
%   uigetfile2.m; ViewpointFigMode{Single, 1, 2, 3, 4}.m; 
%   viewpointMovieFullOneMovie.m; Export_Fig;

%Housekeeping
clc; clear all; close all;
% Common warnings
warning('off','MATLAB:xlsread:ActiveX');
warning('off','MATLAB:mkdir:DirectoryExists');
warning('off','MATLAB:Images:initSize:adjustingMag');
warning('off','MATLAB:getframe:RequestedRectangleExceedsFigureBounds');
warning('off','MATLAB:audiovideo:avifile:indeo5NotFound');
warning('off','MATLAB:audiovideo:VideoWriter:mp4FramePadded');
warning('off','MATLAB:LargeImage');
warning('off','MATLAB:gui:latexsup:UnableToInterpretTeXString');

%% Import the file
[excel, dirname] = uigetfile2('.xlsx');
fullfilename  = strcat(dirname,excel);
[num,txt,raw] = xlsread(fullfilename);

%% 
%assemble the rows
%initialise receiving matrix
vprawData = zeros(size(num,1), 15);

% Column 1 = startDate
startDate = txt(2:size(txt,1),28);%start from 2nd row for numbers only
startDate = str2double(strrep(startDate,'/',''));
vprawData(:,1) = startDate;

% Column 2 = startTime
time = num(:,26);
time = datestr(time, 'HHMM');
startTime = cellstr(time);
%assign to row 2
vprawData(:,2) = str2double(startTime);

% Column 3 = Day (1) or Night (0)
% NightIdx1 = vpData(:,2) < 900; %9am
% NightIdx2 = vpData(:,2) > 2300; %11pm

daynightIdx = (vprawData(:,2) < 900) + (vprawData(:,2) > 2300);
vprawData(:,3) = daynightIdx;

% Column 4 and 5 start and end
startPeriod = vprawData(:,4);
endPeriod = vprawData(:,5);
vprawData(:,4:5) = num(:,12:13);

integrationPeriod = vprawData(1,5);

% Column 6
%extract the %02d number from string
for ii = 1:size(time,1)
    animalNum(ii) = sscanf(raw{ii+1,2},'z%d');
end
vprawData(:,6) = animalNum'; %animal
wellNumber = vprawData(:,6);
totalAnimal = max(vprawData(:,6));
% Column 7-15
vprawData(:,7:15) = num(:,16:24);
freezeCount = vprawData(:,7);
freezeDuration = vprawData(:,8);
midCount = vprawData(:,9);
midDuration = vprawData(:,10);
burstCount = vprawData(:,11);
burstDuration = vprawData(:,12);
zeroCount = vprawData(:,13);
zeroDuration = vprawData(:,14);
activityIntegral = vprawData(:,15);

%% Tidyup recording errors
%if start and end have decimals, delete row.
badframeIdx = round(startPeriod)==startPeriod;
%if there is a decimal value, delete entire row
vprawData(~badframeIdx)=[];
badframeIdx2 = round(endPeriod)==endPeriod;
%if there is a decimal value, delete entire row
vprawData(~badframeIdx2)=[];
%If start and end are the same values delete that row

badframeIdx3 = round(vprawData(:,4))==round(vprawData(:,5));
vprawData(badframeIdx3,:) = [];
badframeIdx4 = round(vprawData(:,4))==(round(vprawData(:,5))-1);
vprawData(badframeIdx4,:) = [];


%% Sort by animal and then start.
%Extract if row has z001 (i) and assign to 3D vector


vpData = zeros(size(vprawData,1)/totalAnimal, 15, totalAnimal);

for jj = 1:totalAnimal
    tempWellIdx = vprawData(:,6) == jj;
    vpData(:,:,jj) = vprawData(tempWellIdx,:);
end

oneDayBool = [];
threehrBool = [];
choice = menu('Is this a 24hr sleep/wake or 3hr drug experiment?','24hr sleep/wake','3hr drug','position','center');
if choice == 1
    oneDayBool = 1;
    threehrBool = 0;
elseif choice == 2 
    oneDayBool = 0;
    threehrBool = 1;
end
%% user defined choices
prompt1 = {'Do you want TotalBoxPlots (1 is yes):',...
            'How many conditions are we comparing?',...
            'Experiment Mode? (0 is single well, 1 is plate, 2 is Row 1.3 vs 2.4,... 3 is Row 1 vs 2 and 3 vs 4, and 4 is Row 1 vs 2.3.4)',...
            'Do you want movies? (1 for yes)','Row 1 condition: '...
            ,'Row 2 condition: ','Row 3 condition: ','Row 4 condition: ',...
            'Do you want to remove any timepoints due to treatment? (1 is yes)',...
            'Do you want to remove the Day/Night transition minute? (1 is yes)',...
            'Set time window Bin width (min): ',...
            'Do you want a sleep/wake matrix? (1 is yes)',...
            'Do you want to exclude any wells? (1 is yes)'};
dlg_title1 = 'Choose User-defined parameters';
num_lines1 = 1;
default1 = {'1','2','2','0','Con','Test','Con','Test','0','1','10','1','0'};
answer1 = inputdlg(prompt1,dlg_title1,num_lines1,default1);
%Assign new values to variables
needTotalBoxBool = str2double(cell2mat(answer1(1)));
rowBool = str2double(cell2mat(answer1(2)));
expMode = str2double(cell2mat(answer1(3)));
moviesBool = str2double(cell2mat(answer1(4)));
rowLabels(1,1) = answer1(5);
rowLabels(1,2) = answer1(6);
rowLabels(1,3) = answer1(7);
rowLabels(1,4) = answer1(8);
treatmentBool = str2double(cell2mat(answer1(9)));
removeTransBool = str2double(cell2mat(answer1(10)));
binWidth = str2double(cell2mat(answer1(11)));
sleepwakeMatrixBool = str2double(cell2mat(answer1(12)));
removeBool = str2double(cell2mat(answer1(13)));
if threehrBool == 1;
    binWidth = 1;
end
%% Exclusion criteria
%before exclusions:
aaa = vpData(:,:,1);

% Remove measurements taken during heatshock or adding drugs.
if(treatmentBool)
    prompt3 = {'Treatment Started: ', 'Treatment Ended: '};
    dlg_title3 = 'Between which datarows did the treatment occur';
    num_lines3 = 1;
    default3 = {'140','200'};
    answer3 = inputdlg(prompt3,dlg_title3,num_lines3,default3);
    startTreat = str2double(cell2mat(answer3(1)));
    endTreat = str2double(cell2mat(answer3(2)));
    vpData(startTreat:endTreat,3:15,:) = NaN;
end
%Exclude the minute bin when the light turns on and off.
%Night -> Day
if(removeTransBool)
    remTransIdx = find(vpData(:,2)==900);
    vpData(remTransIdx,3:15,:) = NaN;
    %Day -> Night
    remTransIdx2 = find(vpData(:,2)==2300);
    vpData(remTransIdx2,3:15,:) = NaN;
end
%% remove any fish that didn't survive the experiment
if(removeBool)
    prompt4 = inputdlg('Enter space-separated numbers:',...
             'Wells to exclude from analysis', [1 75]);
    if ~isempty(prompt4)
        deletedWellIdx = str2num(prompt4{:}); 
        for pp = 1:size(deletedWellIdx,2)
            deleteWell = deletedWellIdx(pp);
            vpData(:,:,deleteWell) = NaN;
        end
    end
end
%% no need to sort!!
noofBins = size(vpData,1)./binWidth;
binTimeIdx = 1:binWidth:size(vpData,1);
binTimeStamp = vpData(binTimeIdx,2,1);
% activity measurement
for kk = 1:size(vpData,3)
    for ll = 1:noofBins
        vpMeasure(ll,1,kk) = vpData((1+binWidth*(ll-1)),2,kk);
        vpMeasure(ll,2,kk) = nansum(vpData(((1+binWidth*(ll-1))):binWidth*ll,10,kk)); %activity
        vpMeasure(ll,3,kk) = nansum(vpData((1+binWidth*(ll-1)):binWidth*ll,10,kk)==0);% sleep
    end
end

%% make sure the nan values of vpData remain nan in vpMeasure:
if(treatmentBool)
    vpMeasureZeroIdx = find(vpMeasure(:,2,1)==0 & vpMeasure(:,3,1)==0); %if both are 0 because of the filters above, make nan
    vpMeasure(vpMeasureZeroIdx,2:3,:) = NaN;
end
if(removeBool)
    if ~isempty(prompt4)
        for qq = 1:size(deletedWellIdx,2)
            deleteWell2 = deletedWellIdx(qq);
            vpMeasure(:,:,deleteWell2) = NaN;
        end
    end
else
    deletedWellIdx = 0;
end
%% Make a folder to store outputs...
mkdir(dirname, 'Summary')
outputdir = strcat(dirname,'Summary',filesep);

%% Reorganise for graphical program
% % All same condition
if any(deletedWellIdx==1)
    vpTime = vpMeasure(:,1,5); %corner case first well is deleted by user
else
    vpTime = vpMeasure(:,1,1); %just use the first column of first sample.
end
aab = vpTime(1);
for mm = 1:size(vpMeasure,3)
    vpWake(:,mm) = vpMeasure(:,2,mm);
end
%export as csv?

for nn = 1:size(vpMeasure,3)
    vpSleep(:,nn) = vpMeasure(:,3,nn);
end
% Time of Day -> Night transition - 11pm
minutePast = mod(vpTime(1,1),10);
d2n = find(vpTime == (2300 + minutePast));

% Time of Night -> Day transition - 9am
n2d = find(vpTime == (900 + minutePast));

%% sleepwakeMatrix
%Produce a grid of awake and sleep actimetry profiles
if(sleepwakeMatrixBool)
    wellMatrix(vpWake, vpSleep, dirname,deletedWellIdx,d2n,n2d,outputdir,threehrBool);
end

%% Boxplot information
vpDaydata = vpData(vpData(:,3)==0,:,:);
vpNightdata = vpData(vpData(:,3)==1,:,:);

if any(deletedWellIdx==1) %corner case where first well is deleted and messes up boxarranged.
    vpDaydata = vpData(vpData(:,3,5)==0,:,:);
    vpNightdata = vpData(vpData(:,3,5)==1,:,:);
end

for zz = 1:size(vpDaydata,3)
    %Column 1 Day Activity
    vpBoxdata(zz,1) = nansum(vpDaydata(:,10,zz))./size(vpDaydata,1); %Activity(sec/min)
    %Column 2 Night Activity
    vpBoxdata(zz,3) = nansum(vpNightdata(:,10,zz))./size(vpNightdata,1); %Activity(sec/min) 
    %Column 3 Day Sleep
    vpBoxdata(zz,2) = (nansum(vpDaydata(:,10,zz)==0)./size(vpDaydata,1))*60; %Sleep(min/hr)
    %Column 4 Night Sleep
    vpBoxdata(zz,4) = (nansum(vpNightdata(:,10,zz)==0)./size(vpNightdata,1))*60; %Sleep(min/hr)
end
%make sure nan propagates through vpBoxdata...
if(removeBool)
    if ~isempty(prompt4)
        for qqq = 1:size(deletedWellIdx,2)
            deleteWell3 = deletedWellIdx(qqq);
            vpBoxdata(deleteWell3,:) = NaN;
        end
    end
end
%% select appropriate experiment mode for output
switch expMode
    case 0
        ViewpointFigModeSingle(dirname,excel,needTotalBoxBool,vpWake,...
        vpSleep,vpTime,vpBoxdata,rowLabels,d2n,n2d,outputdir,threehrBool,...
        oneDayBool);
    case 1
        ViewpointFigMode1(dirname,excel,needTotalBoxBool,vpWake,...
        vpSleep,vpTime,vpBoxdata,rowLabels,d2n,n2d,outputdir,...
        threehrBool,oneDayBool);
    case 2
        ViewpointFigMode2(dirname,excel,needTotalBoxBool,rowBool,vpWake,...
        vpSleep,vpTime,vpBoxdata,rowLabels,...
        d2n,n2d,outputdir,threehrBool,oneDayBool);
    case 3
        ViewpointFigMode3(dirname,excel,needTotalBoxBool,rowBool,vpWake,...
        vpSleep,vpTime,vpBoxdata,rowLabels,d2n,n2d,outputdir,...
        threehrBool,oneDayBool);
    case 4
        ViewpointFigMode4(dirname,excel,needTotalBoxBool,rowBool,vpWake,...
        vpSleep,vpTime,vpBoxdata,rowLabels,d2n,n2d,outputdir,...
        threehrBool,oneDayBool);
end

%%save out
save([outputdir,'plateSupportingData.mat']);
%% Edit and convert Viewpoint's Xvid avi movies
%go through making the videos you want
if (moviesBool)
    while zz == 1;
        viewpointMovieFullOneMovie(excel,dirname,vpTime);
        choice = menu('Do you want to select another ROI?','Yes','No - Im done here!','position','center');
        if choice==1
            viewpointMovieFullOneMovie(excel,dirname,vpTime);
            zz = 1;
        else
           	zz = 0;
        end
    end
end
%clear up
clc; clear all; close all;