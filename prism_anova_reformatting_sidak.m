%{

Prism 2-Way ANOVA Re-Formatting with the Sidak Correction
Molnar Lab 2023
Marissa Mueller

prism_anova_reformatting_sidak.m

%}

%{

Overview: 

This code imports excel sheets outputted as 2-way ANOVA results from
GraphPad Prism. It outputs summary statistics and test details for each 
row factor computed using the Sidak correction (as opposed to the 
two-stage linear setup of Benjamini, Krieger, and Yekutieli).

Code Requirements: 

This script requires the user to specify the input and output directory 
location, as well as additional processing specifications (see Inputs).

Inputs:

 - prismFolderLocation (parent file location for Prism outputs)
 - outputFolderLocation (output file location for re-formatted data)
 - numRF (the number of row factor elements e.g., brain regions)

Outputs:

Re-formatted Prism data, which supplements prism_anova_reformatting

Export and Application:

Outputs are saved in the designated output folder as Name_HEATMAP.csv. 
This file can be exported to a graphing program of choice for further
visualisation and data analysis (i.e., tailored for Prism).

%}

%% Establish working directories and import variables

clear
% Retreive the parent Prism directory
prompt_prismFolderLocation = "Enter the folder path where " + ...
    "parent Prism data are located: ";
prismFolderLocation = input(prompt_prismFolderLocation,"s");
prismFolderLocationChar = convertStringsToChars(prismFolderLocation);
prompt_outputFolderLocation = "Enter the path of the " + ...
    "folder where you would like outputs saved: ";
outputFolderLocation = input(prompt_outputFolderLocation,"s");
% Enter the number of row factor elements (e.g., brain regions) which must
% be the same number for spreadsheets in the input folder. Those with a
% different number of elements must be analysed separately
prompt_numRF = "Enter the number of row factor elements (e.g., " + ...
    "brain regions (7)): ";
numRF = input(prompt_numRF);
% Add folder to the working directory path
addpath(prismFolderLocation,'-end');
% Extract all file names,  
prismFileNames = dir([prismFolderLocationChar, '\*.csv']);
numPrismFiles = length(prismFileNames);

%% Iteratively extract, re-format, and save data

% For each input sheet in the designated folder
for i = 1:numPrismFiles
    % Extract the name of the present Prism sheet
    fileNameHere = prismFileNames(i).name;
    % Import .csv
    prismExcelImport = readcell(fileNameHere);
    % Create or overwrite array to house the names of row factor elements
    nameRFs = strings(numRF,1);
    % Initialise an output matrix to house overall statistical data
    statsOutP = strings(3,numRF);
    % For statsDetails: column 1 = t-stat, column 2 = DF, column 3 = mean
    % difference, and column 4 = SD
    statsDetails = strings(numRF,4);    
    % Start at the top of Prism stats files and work down to extract 1)
    % row factor element names, 2) row factor p-values, 3) Cohen's d, 4) %
    % change, and 5) pairwise row factor p-values resulting from a t-test
    % comparing % changes across column factors. currentRow = 6 to account 
    % for column headers and summary data upon import from Excel to MATLAB
    currentRow = 6;
    % First, extract the names of each row factor element, then  
    % corresponding column factor p-values to the output data matrix. 
    % Adust if needed to ONLY extract the p-values for the effect of a 
    % single column factor 
    for j = 1:numRF
        % Populate an array to house the names of row factor elements
        nameRFs(j,1) = convertCharsToStrings(prismExcelImport{ ...
            (currentRow + j),1});
        % Populating statsOut with individual row factor element p-values
        statsOutP(1,j) = convertCharsToStrings(prismExcelImport{ ...
            (currentRow + j),6});
        % Set to 0.9999 if p > 0.9999
        if statsOutP(1,j) == ">0.9999"
            statsOutP(1,j) = 0.9999;
        end
    end
    % Move currentRow to start at the next block of useful values. Blocks
    % in the equation are separated to 1) skip over the first block of
    % summary information/headers (currentRow), then 2) skip over p-values
    % depending on the number of row factors and column headers (again
    % according to the format once data is imported from Excel to MATLAB)
    currentRow = (currentRow) + (numRF + 2);
    % Scan through rows to calculate Cohen's d and % change for the effect
    % of column factor on each row factor. Adust if needed to ONLY extract 
    % the p-values for the effect of a single column factor
    for j = 1:numRF
        % Extract t-stat
        statsDetails(j,1) = prismExcelImport{(currentRow + j),8};
        % Extract the number of degrees of freedom 
        statsDetails(j,2) = prismExcelImport{(currentRow + j),9};
        % Extract mean 1, mean 2, SE of differences, n1, and n2
        mean1 = prismExcelImport{(currentRow + j),2};
        mean2 = prismExcelImport{(currentRow + j),3};
        diffSE = prismExcelImport{(currentRow + j),5};
        n1 = prismExcelImport{(currentRow + j),6};
        n2 = prismExcelImport{(currentRow + j),7};
        % Calculate n and SD
        pooledN = (n1 + n2)/2;
        diffSDTemp = diffSE*sqrt(pooledN);
        diffSD = round(diffSDTemp,4); 
        % Save SD
        statsDetails(j,4) = num2str(diffSD);
        % Calculate mean difference
        meanDiffTemp = abs(mean2 - mean1);
        meanDiff = round(meanDiffTemp,4);
        % Save mean difference
        statsDetails(j,3) = num2str(meanDiff);
        % Calculate Cohen's d
        cohensDTemp = meanDiff/diffSD;
        cohensD = round(cohensDTemp,4); 
        % Populate statsOut with Cohen's d
        statsOutP(2,j) = num2str(cohensD);
        % Calculate % change
        percentChange = ((mean2 - mean1)/mean1)*100;
        % Populate statsOut with % change
        statsOutP(3,j) = num2str(percentChange);
    end
    % Re-format output table, and begin doing so by initialising a dynamic
    % column index variable
    numColIncrement = numRF;
    if numColIncrement < 6
        numColIncrement = 6;
    end
    % Initialise final output matrix
    finalOutput = strings((5 + numRF),(1 + numColIncrement));
    % Populate row headers for the first data block
    finalOutput(1,1) = "-";
    finalOutput(2,1) = "P-value";
    finalOutput(3,1) = "Cohen's d";
    finalOutput(4,1) = "% Change";
    for n = 1:numRF
        % Populate column headers for the first data block
        finalOutput(1,(1 + n)) = nameRFs(n,1);
        % Populate p-values in the first block
        finalOutput(2,(1 + n)) = statsOutP(1,n);
        % Populate d-values in the first block
        finalOutput(3,(1 + n)) = statsOutP(2,n);
        % Populate % change values in the first block
        finalOutput(4,(1 + n)) = statsOutP(3,n);
    end
    % Populate column headers for the second block
    finalOutput(5,1) = "t (DF) statistic";
    finalOutput(5,1) = "P-value";
    finalOutput(5,1) = "Significance";
    finalOutput(5,1) = "Mean difference (SD)";
    finalOutput(5,1) = "Cohen's d";
    finalOutput(5,1) = "Effect size";
    % Populate data for the second block 
    for k = 1:numRF
        % Populate row headers for the second block
        finalOutput((4 + numRF + 1 - k),1) = nameRFs(k,1);
        % Populate values for the t (DF) statistic
        finalOutput((4 + numRF + 1 - k),2) = "t (" + statsDetails( ...
            k,2) + ") = " + statsDetails(k,1);
        % Populate p-values
        finalOutput((4 + numRF + 1 - k),3) = "P=" + statsOutP(1,k);
        % Populate significance interpretation
        if str2double(statsOutP(1,k)) < 0.0001
            finalOutput((4 + numRF + 1 - k),4) = "****P<0.0001";
        elseif str2double(statsOutP(1,k)) < 0.001
            finalOutput((4 + numRF + 1 - k),4) = "***P<0.001";
        elseif str2double(statsOutP(1,k)) < 0.01
            finalOutput((4 + numRF + 1 - k),4) = "**P<0.01";
        elseif str2double(statsOutP(1,k)) < 0.05
            finalOutput((4 + numRF + 1 - k),4) = "*P<0.05";
        else
            finalOutput((4 + numRF + 1 - k),4) = "NS";
        end
        % Populate mean difference (SD)
        finalOutput((4 + numRF + 1 - k),5) = statsDetails( ...
            k,3) + " (" + statsDetails(k,4) + ")";
        % Populate d-values
        finalOutput((4 + numRF + 1 - k),6) = "d=" + statsOutP(2,k);
        % Populate effect size interpretation
        if str2double(statsOutP(2,k)) > 0.8
            finalOutput((4 + numRF + 1 - k),7) = "Large";
        elseif str2double(statsOutP(2,k)) > 0.5
            finalOutput((4 + numRF + 1 - k),7) = "Moderate";
        elseif str2double(statsOutP(2,k)) > 0.2
            finalOutput((4 + numRF + 1 - k),7) = "Small";
        else
            finalOutput((4 + numRF + 1 - k),7) = "Very small";
        end
    end
    % Save output table, first creating an output folder if one does not
    % already exist
    heatmapFolder = "QUINT-Prism_Heat-Map_Data_SIDAK-ALL";
    outputDataFolderLocation = outputFolderLocation + "\" + heatmapFolder;
    % Make a new folder where outputs will be saved, if the folder does 
    % not already exist, and add it to the working directory
    if isfolder(heatmapFolder) == 0
       mkdir(outputFolderLocation,heatmapFolder)
       % Add this path to the working directory
       addpath(heatmapFolder,'-end');
    else
        addpath(heatmapFolder,'-end');
    end
    % Name output table
    fileNameHereNoExtension = extractBefore(fileNameHere, ".");
    saveNameHeatMap = fileNameHereNoExtension + "_HEATMAP.csv";
    savePathHeatMap = outputDataFolderLocation + "\" + saveNameHeatMap;
    % Save file
    writematrix(finalOutput,savePathHeatMap)
    % Repeat for each file
end
fprintf("Program complete.\n")