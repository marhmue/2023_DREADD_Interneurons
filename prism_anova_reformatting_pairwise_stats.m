%{

Prism 2-Way ANOVA Re-Formatting for 2-Stage Pairwise Comparisons
Molnar Lab 2023
Marissa Mueller

prism_anova_reformatting_pairwise_stats.m

%}

%{

Overview: 

This code imports excel sheets outputted as 2-way ANOVA results from
GraphPad Prism. It outputs summary statistics and test details for each 
row factor computed using the two-stage linear setup of Benjamini, 
Krieger, and Yekutieli (as opposed to the Sidak correction).

Code Requirements: 

This script requires the user to specify the input and output directory 
location, as well as additional processing specifications (see Inputs).

Inputs:

 - prismFolderLocation (parent file location for Prism outputs)
 - outputFolderLocation (output file location for re-formatted data)
 - numRF (the number of row factor elements e.g., 7 brain regions)
 - numCF (the number of column factor elements e.g., 2 CNO conditions)

Outputs:

Re-formatted Prism data. Rows and columns of the final output sheet 
constitute row factors (brain regions of interest, which are extracted 
from the input sheet). The first row atop the main map houses % change  
values (column factor, for example saline vs. CNO). The row above houses 
Cohen's d values representing the effect size of the column factor for 
each row factor (brain region). The top row houses p-values as a measure
of significance for the column factor (saline vs CNO) on the row factor 
(for each brain region). These p-values are extracted as direct outputs of
multiple comparisons in the Prism 2-way ANOVA. Cohen's d values are
calculated by the standard formula for effect size. % change is calculated
as the absolute value of the mean difference over the pooled standard
deviation for the column factor (saline vs CNO). The main body of the heat
map contains p-values comparing the mean differences across row factors
(brain regions). That is, it represents the p-value resulting from a 
simple t-test comparing the mean difference between column factors (saline
vs. CNO) between and across row factors (brain regions). These populate
the upper right triangular portion, and corresponding d-values (calculated
again using the standard formula for effect size) populate the lower left
triangular portion of square matrices for subsequent import to Prism heat 
map templates.

Export and Application:

Outputs are saved in the parent folder as _HEATMAP_2.csv. This  
file can be exported to a graphing program of choice for further
visualisation and data analysis (e.g., back to Prism).

%}

%% Establish working directories and import variables

clear
% Retreive the parent Prism directory
prompt_prismFolderLocation = "Enter the folder path where " + ...
    "parent Prism data are located: ";
prismFolderLocation = input(prompt_prismFolderLocation,"s");
prismFolderLocationChar = convertStringsToChars(prismFolderLocation);
% Determine the output directory
prompt_outputFolderLocation = "Enter the path of the " + ...
    "folder where you would like outputs saved: ";
outputFolderLocation = input(prompt_outputFolderLocation,"s");
% Enter the number of row factor elements (e.g., brain regions) which must
% be the same number for spreadsheets in the input folder. Those with a
% different number of elements must be analysed separately
prompt_numRF = "Enter the number of row factor elements (e.g., " + ...
    "brain regions): ";
numRF = input(prompt_numRF);
% Enter the name of column factor elements (e.g., saline vs. CNO)
prompt_numCF = "Enter the number of column factor elements (e.g., " + ...
    "saline vs. CNO): ";
numCF = input(prompt_numCF);
% Add folder to the working directory path
addpath(prismFolderLocation,'-end');
% Extract all file names,  
prismFileNames = dir([prismFolderLocationChar, '\*.csv']);
numPrismFiles = length(prismFileNames);
% Iterating through for each input sheet
for i = 1:numPrismFiles
    % Extract the name of the present Prism sheet
    fileNameHere = prismFileNames(i).name;
    % Import .csv
    prismExcelImport = readcell(fileNameHere);
    % Create or overwrite array to house the names of row factor elements
    nameRFs = strings(numRF,1);
    % Initialise a statistics output matrix according to the format
    % specified in 'Outputs', one though for p-values in the diagonals and
    % one for d-values in the diagonals
    statsOutP = strings((3 + numRF),numRF);
    statsOutD = strings((3 + numRF),numRF);
    % Determine the number of combinations of column factor elements
    numCFCombos = numCF*(numCF - 1)/2;
    % Starting at the top of the stats file and working down to extract 1)
    % row factor element names, 2) row factor p-values, 3) Cohen's d, 4) %
    % change, and 5) pairwise row factor p-values resulting from a t-test
    % comparing % change across column factors. Starting at 7 to account
    % for summary data and column headers
    currentRow = 6;
    % First extract the names of each row factor element, then extract the
    % corresponding column factor p-values to the output data matrix. Adust
    % if needed to ONLY extract the p-values for the effect of a single
    % column factor
    for j = 1:numRF
        % Populate an array to house the names of row factor elements
        nameRFs(j,1) = convertCharsToStrings(prismExcelImport{( ...
            currentRow + 1 + (1 + numCFCombos)*(j - 1)),1});
        % Populating statsOut with individual row factor element p-values
        statsOutP(1,j) = convertCharsToStrings(prismExcelImport{( ...
            currentRow + 1 + (1 + numCFCombos)*(j - 1) + 1),5});
        % Set to 0.9999 if p > 0.9999
        if statsOutP(1,j) == ">0.9999"
            statsOutP(1,j) = 0.9999;
            %statsOutD(1,j) = 0.9999;
        end
    end
    % Move currentRow to start at the next block of useful values. Blocks
    % in the equation are separated to 1) skip over the first block of
    % summary information/headers (currentRow), 2) skip over the first set
    % of p-values depending on the number of row and column factors, 3) 
    % skip over p-values for combinations of row factors for each column 
    % factor, then 4) adding 1 to account for the section header 'Test
    % Details' from which the next set of statistics will be calculated
    currentRow = (currentRow) + (numRF*((numCF*(numCF - 1)/2) ...
        + 1)) + ((1 + ((numRF*(numRF - 1)/2)))*numCF) + (1);
    % Scan through rows to calculate Cohen's d and % change for the effect
    % of column factor on each row factor. Adust if needed to ONLY extract 
    % the p-values for the effect of a single column factor
    for j = 1:numRF
        % Scan to the row with data for the column factor of interest
        CFofInterest = 2;
        % Adjust CFofInterest if there are more than 2 column factors and
        % to therefore select the one presently of interest
        currentRow = currentRow + CFofInterest;
        % Extract mean 1, mean 2, SE of differences, n1, and n2
        mean1 = prismExcelImport{currentRow,2};
        mean2 = prismExcelImport{currentRow,3};
        diffSE = prismExcelImport{currentRow,5};
        n1 = prismExcelImport{currentRow,6};
        n2 = prismExcelImport{currentRow,7};
        % Calculate Cohen's d
        pooledN = (n1 + n2)/2;
        diffSD = diffSE*sqrt(pooledN);
        cohensD = abs(mean2 - mean1)/diffSD;
        % Populate statsOut with Cohen's d
        statsOutP(2,j) = num2str(cohensD);
        statsOutD(2,j) = num2str(cohensD);
        % Calculate % change
        percentChange = ((mean2 - mean1)/mean1)*100;
        % Populate statsOut with % change
        statsOutP(3,j) = num2str(percentChange);
        statsOutD(3,j) = num2str(percentChange);
    end
    % Move to the next set of blocks from which data will be extracted to
    % populate the main body of the heat map. This will involve performing
    % an unpaired t-test on the differences between the effect of the
    % selected column factor between all pairs of row factors. First 
    % select the column factors to be considered. 1 and 2 are used here,
    % but can be adjusted if there are more than two column factors
    cf1ID = 1;
    cf2ID = 2;
    % Determine the number of rows separating pairs of column factor data
    % for the same row factor combination
    spacer = (1 + ((numRF*(numRF - 1)/2)));
    % Find the starting row for the paired cf being considered
    pairedRow = currentRow + 1 + spacer*(cf2ID - 1);
    % Increasing currentRow to skip over the section header. The second
    % section header is accounted for in spacer
    currentRow = currentRow + 1 + spacer*(cf1ID - 1);
    % Initialise variable to track the highest recorded Cohen's d value
    highestCohensD = 0;
    % Initialise index counters to determine the rows and columns to be
    % populated in statsOut
    pairwiseIncrementer = 1;
    rowSOQUR = pairwiseIncrementer;
    colSOQUR = pairwiseIncrementer;
    rowSOQLL = pairwiseIncrementer;
    colSOQLL = pairwiseIncrementer;
    % Generate a list for each of the key pairwise statistics variables
    listHeight = 0;
    for h = 1:(numRF - 1)
        listHeight = listHeight + h;
    end
    listCOMPARISON = strings(listHeight,1);
    listTSTAT = strings(listHeight,1);
    listDF = strings(listHeight,1);
    listPVALUE = strings(listHeight,1);
    listSIG = strings(listHeight,1);
    listMEANDIFF = strings(listHeight,1);
    listSD = strings(listHeight,1);
    listCOHENSD = strings(listHeight,1);
    listEFFECTSIZE = strings(listHeight,1);
    % Scanning through each row factor pairing combinatorically, starting
    % at all options for the first element and ending at the final single
    % unaccounted for pairing of the last two elements
    for j = 1:((numRF*(numRF - 1)/2))
        % Row for cf1 data
        rowCF1Here = currentRow + j;
        % Row for cf2 data
        rowCF2Here = pairedRow + j;
        % Extract the name of the comparison being performed, which 
        % applies to both CF1 and CF2
        listCOMPARISON(j,1) = prismExcelImport{rowCF1Here,1};
        % Remove "C" to match heat maps
        listCOMPARISON(j,1) = erase( listCOMPARISON(j,1) , "C" );
        % Extract mean, SE of differences, n1, and n2 for CF1
        meanDiffCF1 = prismExcelImport{rowCF1Here,4};
        meanDiffSECF1 = prismExcelImport{rowCF1Here,5};
        n1CF1 = prismExcelImport{rowCF1Here,6};
        n2CF1 = prismExcelImport{rowCF1Here,7};
        % Calculate pooled n and SD of differences for CF1
        pooledNCF1 = (n1CF1 + n2CF1)/2;
        diffSDCF1 = meanDiffSECF1*sqrt(pooledNCF1);
        % Extract mean, SE of differences, n1, and n2 for CF2
        meanDiffCF2 = prismExcelImport{rowCF2Here,4};
        meanDiffSECF2 = prismExcelImport{rowCF2Here,5};
        n1CF2 = prismExcelImport{rowCF2Here,6};
        n2CF2 = prismExcelImport{rowCF2Here,7};
        % Calculate pooled n and SD of differences for CF2
        pooledNCF2 = (n1CF2 + n2CF2)/2;
        diffSDCF2 = meanDiffSECF1*sqrt(pooledNCF2);
        % Calculate t statistic precursor values assuming equal variances
        tStatNumerator = meanDiffCF2 - meanDiffCF1;
        tStatDenominatorPooledSD = sqrt((((pooledNCF2 - 1)*( ...
            diffSDCF2^2)) + ((pooledNCF1 - 1)*(diffSDCF1^2)))/( ...
            pooledNCF2 + pooledNCF1 - 2));
        % Calculate t-statistic
        tStatDenominatorAddOn = sqrt((1/pooledNCF2) + (1/pooledNCF1));
        tStat = tStatNumerator/(tStatDenominatorPooledSD* ...
            tStatDenominatorAddOn);
        % Calculate DF
        dfHere = pooledNCF2 + pooledNCF1 - 2; 
        % Calculate p-value 
        pHere = tcdf(tStat,dfHere);
        % Calculate Cohen's d
        dHere = abs(tStatNumerator)/tStatDenominatorPooledSD;
        if dHere > highestCohensD
            highestCohensD = dHere;
        end
        % Determine the row/column to be populated in QUR of statsOut,
        % which will be populated with p-values
        colSOQUR = colSOQUR + 1;
        if colSOQUR > numRF
            pairwiseIncrementer = pairwiseIncrementer + 1;
            colSOQUR = pairwiseIncrementer + 1;
        end
        rowSOQUR = 3 + pairwiseIncrementer;
        % Determine the row/column to be populated in ALL of statsOut,
        % which will be populated with d-values
        colSOQLL = pairwiseIncrementer;
        rowSOQLL = 3 + colSOQUR;
        % Populate the upper right quadrant of statsOut with the p-value
        statsOutP(rowSOQUR,colSOQUR) = num2str(pHere);
        if statsOutP(rowSOQUR,colSOQUR) == ">0.9999"
            statsOutP(rowSOQUR,colSOQUR) = 0.9999;
        end
        % Populate the lower left quadrant of statsOut with Cohen's d
        statsOutD(rowSOQLL,colSOQLL) = num2str(dHere);
        % Save mean differences
        listMEANDIFF(j,1) = num2str(tStatNumerator);
        % Save SD
        listSD(j,1) = num2str(tStatDenominatorPooledSD);
        % Save t-statistic
        listTSTAT(j,1) = num2str(tStat);
        % Save DF
        listDF(j,1) = num2str(dfHere);
        % Populate the p-value list
        listPVALUE(j,1) = num2str(pHere);
        % Populate the significance interpretation list
        if pHere < 0.0001
            listSIG(j,1) = "****P<0.0001";
        elseif pHere < 0.001
            listSIG(j,1) = "***P<0.001";
        elseif pHere < 0.01
            listSIG(j,1) = "**P<0.01";
        elseif pHere < 0.05
            listSIG(j,1) = "*P<0.05";
        else
            listSIG(j,1) = "NS";
        end
        % Populate the d-value list
        listCOHENSD(j,1) = num2str(dHere);
        % Populate the effect size interpretation list
        if dHere > 0.8
            listEFFECTSIZE(j,1) = "Large";
        elseif dHere > 0.5
            listEFFECTSIZE(j,1) = "Moderate";
        elseif dHere > 0.2
            listEFFECTSIZE(j,1) = "Small";
        else
            listEFFECTSIZE(j,1) = "Very small";
        end
        % Repeat for each row factor combination, leaving the diagonal
        % populated with ones
    end
    % Populating diagonals of statsOut
    for s = 1:numRF
        statsOutP((3 + s),s) = "1";
        statsOutD((3 + s),s) = "0";
    end
    % Re-format output table
    numColIncrement = numRF;
    if numColIncrement < 7
        numColIncrement = 7;
    end
    finalOutput = strings((5 + numRF + 1 + numRF + 2 + ((numRF*( ...
        numRF - 1)/2))),(1 + numColIncrement));
    % Populate static headers
    finalOutput(1,1) = "-";
    finalOutput(2,1) = "P-value";
    finalOutput(3,1) = "Cohen's d";
    finalOutput(4,1) = "% Change";
    finalOutput(5,1) = "Pairwise P-Value";
    finalOutput(5 + numRF + 1,1) = "Pairwise Cohen's d";
    finalOutput(5 + numRF + 1 + numRF + 1,1) = "Statistics Details";
    for n = 1:numColIncrement
        % Populate dynamic column headers
        finalOutput(1,(1 + n)) = nameRFs(n,1);
        finalOutput(5,(1 + n)) = "-";
        finalOutput((5 + numRF + 1),(1 + n)) = "-";
        % Populate dynamic row headers
        finalOutput((5 + n),1) = nameRFs(n,1);
        finalOutput((5 + numRF + 1 + n),1) = nameRFs(n,1);
        % Populate p-values
        finalOutput(2,(1 + n)) = statsOutP(1,n);
        % Populate d-values
        finalOutput(3,(1 + n)) = statsOutP(2,n);
        % Populate % change values
        finalOutput(4,(1 + n)) = statsOutP(3,n);
        finalOutput((5 + numRF + 1 + numRF + 1),(1 + n)) = "-";
        % Populate statistics details column headers
        finalOutput((5 + numRF + 1 + numRF + 1 + 1),1) = "Comparison";
        finalOutput((5 + numRF + 1 + numRF + 1 + 1),2) = "t (DF) " + ...
            "statistic";
        finalOutput((5 + numRF + 1 + numRF + 1 + 1),3) = "P-value";
        finalOutput((5 + numRF + 1 + numRF + 1 + 1),4) = "Significance";
        finalOutput((5 + numRF + 1 + numRF + 1 + 1),5) = "Mean " + ...
            "difference (SD)";
        finalOutput((5 + numRF + 1 + numRF + 1 + 1),6) = "Cohen's d";
        finalOutput((5 + numRF + 1 + numRF + 1 + 1),7) = "Effect size";
        % Populate pairwise p- and d-values
        for m = 1:numRF
            if statsOutP((3 + n),m) ~= ""
                finalOutput((5 + n),(1 + m)) = statsOutP((3 + n),m);
            end
            if statsOutD((3 + n),m) ~= ""
                finalOutput((5 + numRF + 1 + n),(1 + m ...
                    )) = statsOutD((3 + n),m);
            end
        end
    end
    % Populate statistics details
    for r = 1:((numRF*(numRF - 1)/2))
        % Name of comparison
        finalOutput(5 + numRF + 1 + numRF + 1 + 1 + r,1 ...
            ) = listCOMPARISON(r,1);
        % t (DF) statistic
        finalOutput(5 + numRF + 1 + numRF + 1 + 1 + r,2 ...
            ) = "t (" + listDF(r,1) + ") = " + listTSTAT(r,1);
        % P-value
        finalOutput(5 + numRF + 1 + numRF + 1 + 1 + r,3 ...
            ) = "P=" + listPVALUE(r,1);
        % Significance
        finalOutput(5 + numRF + 1 + numRF + 1 + 1 + r,4) = listSIG(r,1);
        % Mean difference (SD)
        finalOutput(5 + numRF + 1 + numRF + 1 + 1 + r,5 ...
            ) = listMEANDIFF(r,1) + " (" + listSD(r,1) + ")";
        % Cohen's d
        finalOutput(5 + numRF + 1 + numRF + 1 + 1 + r,6 ...
            ) = "d=" + listCOHENSD(r,1);
        % Effect size
        finalOutput(5 + numRF + 1 + numRF + 1 + 1 + r,7 ...
            ) = listEFFECTSIZE(r,1);
    end
    % Save output table, first creating an output folder if one does not
    % already exist
    heatmapFolder = "QUINT-Prism_Heat-Map_Data_STATS";
    outputDataFolderLocation = outputFolderLocation + "\" + heatmapFolder;
    % Make a new folder where outputs will be saved, if the folder does 
    % not already exist, and add it to the working directory
    if isfolder(heatmapFolder) == 0
       mkdir(outputFolderLocation,heatmapFolder)
       % Add this path to the working directory
       addpath(heatmapFolder,'-end');
    end
    % Name output table
    fileNameHereNoExtension = extractBefore(fileNameHere, ".");
    saveNameHeatMap = fileNameHereNoExtension + "_HEATMAP_2.csv";
    savePathHeatMap = outputDataFolderLocation + "\" + saveNameHeatMap;
    % Save file
    writematrix(finalOutput,savePathHeatMap)
    % Repeat for each file
end
fprintf("Program complete.\n")