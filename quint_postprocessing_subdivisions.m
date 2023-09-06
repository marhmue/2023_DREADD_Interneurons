%{

QUINT Data Processing
Molnar Lab 2023
Marissa Mueller

quint_postprocessing_subdivisions.m

%}

%{

Overview: 

This code imports QUINT pipeline output data files, re-formatting sheets
for data visualisation and analysis in GraphPAD PRISM. It can be tailored 
to return either cell count and density information for all brain regions 
of interest.

Code Requirements: 

This script requires that users provide the location of QUINT output 
directories and user-inputted processing specifications (see Inputs).

Inputs:

 - quintFolderLocation (parent file location for QUINT outputs)
 - name.txt (SnackerTracker SD output file)
 - numCons (the number of conditions/genotypes being assessed)
 - nameCons (name each condition/genotype)
 - numExpReps (number of experimental replicates/animals per condition)
 - nameExpReps (name each experimental replicate/animal)
 - numTechReps (the number of technical replicates/sections per 
   experimental replicate/animal)
 - numBrainRegions (the number of brain regions assessed by QUINT outputs)

Outputs:

Processed and formatted QUINT output data.

Export and Application:

Outputs are saved in the parent folder as _Density.csv. This file can be 
exported to a graphing program of choice (i.e., Prism) for further
visualisation and data analysis. Columns are grouped by brain region, with
condition/genotype status being discriminated by row. Experimental
replicates (animals) exist as sub-columns with technical replicates 
(repeat slices) collapsed into single averaged measurements.

%}

%% Establish working directories and import variables

clear
entriesOK = "N";
while entriesOK ~= "Y"
    % Retreive the parent QUINT directory for future navigation
    prompt_quintFolderLocation = "Enter the folder path for the " + ...
        "parent QUINT data location: ";
    quintFolderLocation = input(prompt_quintFolderLocation,"s");
    quintFolderLocationChar = convertStringsToChars(quintFolderLocation);
    % Retreive the name of the input Excel sheet for analysis
    prompt_quintExcelFileName = "Enter the name of the input Excel " + ...
        "sheet being used for analysis: ";
    excelFileName = input(prompt_quintExcelFileName,"s");
    excelFileNameChar = convertStringsToChars(excelFileName);
    % Error-checking to ensure the folder and input sheet directory 
    % locations suffice and that the excel data file is not duplicated 
    % in the path specification
    fileNameContainLength = length(excelFileNameChar) - 1;
    fileNameIndexLength = length(excelFileNameChar) + 1;
    % If the folder/file path is long enough to contain fileName
    if length(quintFolderLocationChar) > fileNameContainLength
        % If fileName is included in the user-defined path
        if convertCharsToStrings(quintFolderLocationChar(( ...
                end-fileNameContainLength):end)) == excelFileName
            % Truncate excelFileName and convert data type
            quintFolderLocationChar = quintFolderLocationChar( ...
                1:(end-fileNameIndexLength));
            quintFolderLocation = convertCharsToStrings( ...
                quintFolderLocationChar);
            % quintExcelFileLocation simply = quintFolderLocation
            quintExcelFileLocation = quintFolderLocation;
            quintExcelFileLocationChar = quintFolderLocationChar;
        end
    end
    % If excelFileName was not included in the user-defined path
    if convertCharsToStrings(quintFolderLocationChar(( ...
            end-fileNameContainLength):end)) ~= excelFileName
        % The folder string is already defined as quintFolderLocation
        % Append fileName
        quintExcelFileLocation = quintFolderLocation + "\" + ...
            "" + excelFileName;
        quintExcelFileLocationChar = convertStringsToChars( ...
            quintExcelFileLocation);
    end
    % Extract data from the input excel sheet after first adding the 
    % skrtkdFileLocation folder to the working directory path
    addpath(quintFolderLocation,'-end');
    % Import QUINT-Analy.txt
    fprintf("Importing information from ");
    disp(quintExcelFileLocation);
    % Extract excelFileName.txt
    quintExcelImport = readcell(excelFileNameChar);
    % Extract individual inputs, first the number of conditions/genotypes 
    % for comparison 
    numCons = quintExcelImport{3,6};
    % Number of experimental replicates across all conditions/genotypes
    maxNumExpReps = quintExcelImport{4,6};
    % Maximum number of technical replicates across all conditions/
    % genotypes and experimental replicates
    maxNumTechReps = quintExcelImport{5,6};
    % Initialise output data scaffold array
    numColumns = 5;
    numRows = 1*numCons*maxNumExpReps;
    dataScaffoldStr = strings(numRows,2);
    dataScaffoldNums = zeros(numRows,3);
    dataScaffoldNums(1,1) = numCons;
    % Initialise array to house the names of each experimental
    % condition/genotype for later reference
    nameCons = strings(numCons,1);
    % For each condition/genotype
    for i = 1:numCons
        % Assign a name to each experimental condition/genotype
        nameConHere = convertCharsToStrings(quintExcelImport{( ...
            11 + (i - 1)*maxNumExpReps),1});
        % Store in an array for later reference
        nameCons(i,1) = nameConHere;
        conRow = 1 + (i - 1)*maxNumExpReps;
        dataScaffoldStr(conRow,1) = nameConHere;
        % Number of experimental replicates for this condition
        numExpReps = quintExcelImport{(11 + (i - 1)*maxNumExpReps),2};
        dataScaffoldNums(conRow,2) = numExpReps;
        % For each experimental replicate
        for j = 1:numExpReps
            % Extract the name/ID of each experimental replicate
            nameExpReps = convertCharsToStrings(quintExcelImport{( ...
                11 + (i - 1)*maxNumExpReps + (j - 1)),3});
            nameExpRepsRow = conRow + (j - 1);
            dataScaffoldStr(nameExpRepsRow,2) = nameExpReps;
            % Extract the number of technical replicates for this 
            % condition and experimental replicate
            numTechReps = quintExcelImport{(11 + (i - 1 ...
                )*maxNumExpReps + (j - 1)),4};
            dataScaffoldNums(nameExpRepsRow,3) = numTechReps;
        end
    end
    % Extract the number of brain regions
    numBrainRegions = quintExcelImport{6,6};
    % Extract the number of fluorophore combinations
    numFluoroCombos = quintExcelImport{7,6};
    fluNameHere = strings(numFluoroCombos,1);
    for g = 1:numFluoroCombos
        % Extract the name of each fluorophore combination
        fluNameHere(g,1) = convertCharsToStrings( ...
            quintExcelImport{(11 + (g - 1)),6});
    end
    % Extract the the area (microns, um^2) corresponding to one pixel
    squareMicronsPP = quintExcelImport{8,6};
    fprintf("Data successfully imported.");
    % Display inputted data to ensure entries are correct
    fprintf("\n");
    fprintf("Replicate information:\n" );
    disp(dataScaffoldStr);
    disp(dataScaffoldNums);
    fprintf("\n");
    fprintf("Number of Brain Regions = " + numBrainRegions);
    fprintf("\n");
    fprintf("Fluorophore Combinations: \n");
    for s = 1:numFluoroCombos
        fprintf("# " + num2str(s) + " = " + fluNameHere(s,1));
        fprintf("\n");
    end
    fprintf("Microns per pixel = " + squareMicronsPP ...
        + " um^2/px");
    fprintf("\n");
    % Ask whether these are correct (i.e., whether the program should 
    % proceed or whether inputs need to be re-entered)
    fprintf("Are these entries correct?\n");
    prompt_entriesOK = "Enter 'Y' if 'Yes', or 'N' if 'No' " + ...
        "to re-enter values: ";
    entriesOK = input(prompt_entriesOK,"s");
    % Error-checking for a valid entry
    entriesOK = quint_isentryerror("Y","N",entriesOK);
end
fprintf("Entries confirmed.\n");
fprintf("\n");

%% Data analysis

% Navigate directories and sequentially extract QUINT output data for each
% fluorophore combination of interest
for f = 1:numFluoroCombos
    % Initialise an array for final outputs. Columns will represent data
    % from experimental replicates per brain region. Rows will represent
    % information retrieved for each experimental condition/genotype
    numColumnsFinal = numBrainRegions*maxNumExpReps;
    numRowsFinal = numCons;
    finalOutputTableCellCounts = ones(numRowsFinal,numColumnsFinal).*1000;
    finalOutputTableDensity = ones(numRowsFinal,numColumnsFinal).*1000;
    % Iterate through directories to extract QUINT outputs
    % For each experimental condition/genotype in the following order:
    % EXCITATORY, SALINE
    % EXCITATORY, CNO
    % INHIBITORY, SALINE
    % INHIBITORY, CNO
    % This order also applies to the block of summary data beneath, with
    % the first three rows being EX-Sal, the next three Ex-CNO, the next
    % three INH-Sal, and the final three INH-CNO
    for i = 1:dataScaffoldNums(1,1)
        % Store the index of the first row for this genotype/condition 
        baseRow = 1 + (i - 1)*maxNumExpReps;
        % Repeating for each experimental replicate
        for j = 1:dataScaffoldNums(baseRow,2)
            % Store the index of the row corresponding to the present 
            % experimental replicate for this genotype/condition 
            expRepRow = baseRow + (j - 1);    
            % Extract the name of the present experimental replicate
            expRepHere = dataScaffoldStr(expRepRow,2);
            % Identify the folder path of Nutil outputs given a consistent
            % directory scaffold
            intFPath = quintFolderLocation + "\" + expRepHere + "" + ...
                "\8_Nutil_Output\" + fluNameHere(f,1) + "\Reports\";
            % Retrieve folder contents
            d = dir(convertStringsToChars(intFPath));
            % Remove all files (isdir property is 0)
            dfolders = d([d(:).isdir]);
            % Remove '.' and '..' to only extract the names of sub-folders
            dfolders = dfolders(~ismember({dfolders(:).name},{'.','..'}));
            fileFolder = "";
            % For every sub-folder
            for m = 1:length(dfolders)
                folderNameHere = dfolders(m).name;
                folderNameHereStr = convertCharsToStrings(folderNameHere);
                % Find the folder which contains spreadsheets for custom
                % regions
                if contains(folderNameHereStr,"CustomRegions") == 1
                    % Assign this folder name for future access
                    fileFolder = folderNameHereStr;
                end
            end
            % Append this folder name to folderLocationPath
            folderLocationPath = intFPath + fileFolder;
            % Add this folder to the working directory 
            addpath(folderLocationPath,'-end');
            % Repeating for each technical replicate for the present
            % experimental replicate for cell counts and densities across
            % brain regions
            rawTRHCellCounts = zeros( ...
                maxNumTechReps,numBrainRegions);
            rawTRHDensity = zeros( ...
                maxNumTechReps,numBrainRegions);
            avgTechRepsCellCounts = zeros(1,numBrainRegions);
            avgTechRepsDensity = zeros(1,numBrainRegions);
            % For each technical replicate for the present experimental
            % replicate
            for n = 1:dataScaffoldNums(expRepRow,3)
                % Specify the standardised file name according to QUINT
                % formatting
                num = n;
                padNumHere = num2str(num,'%03.f');
                fileName = fileFolder + "__s" + padNumHere;
                fileNameChar = convertStringsToChars(fileName);
                % Append fileName to the folder path
                fileLocationPath = folderLocationPath + "\" + fileName;
                fprintf("Importing information from ");
                disp(fileLocationPath);
                % Extract information from .csv file
                quintImport = readcell(fileNameChar);
                % If it is the first iteration
                if i == 1 && j == 1
                    % Create array to house the names of each brain region
                    nameBrainRegions = strings(1,numBrainRegions);
                    % For each brain region
                    for q = 2:(numBrainRegions + 1)
                        % Extract the name of the brain region
                        nameBrainRegions(1,(q - 1) ...
                            ) = convertCharsToStrings(quintImport{q,1});
                    end
                end
                % For each brain region, starting at row p = 2 to account 
                % for column headers
                for p = 2:(numBrainRegions + 1)
                    % Preserving the order of cortical regions, and
                    % extracting data according to experimental/technical
                    % replicate status and corresponding location in
                    % finalOutputTable
                    numObjectsImport = quintImport(p,5);
                    % Convert to numerical format
                    numObjectsHere = cell2mat(numObjectsImport);
                    % Simply return cell counts (later). Note that, for 
                    % cortical layer assessments, this will return the
                    % number of cells per cortical layer in a format which
                    % is compatible for direct copy-paste into Prism and
                    % specific to grouped table formatting. Data can be
                    % visualised nicely as a stacked bar plot
                    regionPixelsImport = quintImport(p,2);
                    % Convert to numerical format
                    regionPixelsHere = cell2mat(regionPixelsImport);
                    % Calculate density
                    regionAreaHere = regionPixelsHere*squareMicronsPP;
                    densityHere = numObjectsHere/regionAreaHere;
                    % Only populate rawTechRepsHolder with the calculated
                    % density/cell count value if a valid result is 
                    % returned (that is, it is NOT "NaN"
                    if num2str(densityHere) ~= "NaN"
                        rawTRHCellCounts(n,(p - 1 ...
                            )) = numObjectsHere;
                        rawTRHDensity(n,(p - 1 ...
                            )) = densityHere;
                    end
                end
            end
            % Calculating the average value across technical replicates 
            % for the present experimental replicate
            for q = 1:numBrainRegions
                sumHereCounts = 0;
                sumHereDensity = 0;
                for r = 1:dataScaffoldNums(expRepRow,3)
                    sumHereCounts = sumHereCounts + rawTRHCellCounts(r,q);
                    sumHereDensity = sumHereDensity + rawTRHDensity(r,q);
                end
                avgTechRepsCellCounts(1,q ...
                    ) = sumHereCounts/dataScaffoldNums(expRepRow,3);
                avgTechRepsDensity(1,q ...
                    ) = sumHereDensity/dataScaffoldNums(expRepRow,3);
                % Populate corresponding rows and columns of final output
                % tables with the average of technical replicates
                if num2str(avgTechRepsDensity(1,q)) ~= "NaN"
                    finalOutputTableCellCounts(i,(j + (q - 1 ...
                        )*maxNumExpReps)) = avgTechRepsCellCounts(1,q);
                    finalOutputTableDensity(i,(j + (q - 1 ...
                        )*maxNumExpReps)) = avgTechRepsDensity(1,q);
                end
            end
        end
    end
    
    %% Re-format output table (featuring easy import to Prism)
    
    % For the number of rows, include (in the order of appearance in the
    % equaltion) a column header, a data row for each experimental
    % condition, a blank 'spacer' row, and three rows (in the first column 
    % of each brain region) per experimental condition for the mean, SEM, 
    % and n of technical replicates. This last set of rows is constructed
    % for compatibility with GraphPad prism and ease of data transfer
    finalOutputTableStrCellCounts = strings((1 + height( ...
        finalOutputTableCellCounts) + 1 + numCons*3),width( ...
        finalOutputTableCellCounts));
    finalOutputTableStrDensity = strings((1 + height( ...
        finalOutputTableCellCounts) + 1 + numCons*3),width( ...
        finalOutputTableDensity));
    for i = 1:(height(finalOutputTableDensity) + 2)
        % For the first row, simply populate column headers with the name
        % of each brain region
        if i == 1 
            for k = 1:numBrainRegions
                finalOutputTableStrCellCounts(1,(1 + (k - 1 ...
                    )*maxNumExpReps)) = nameBrainRegions(1,k);
                finalOutputTableStrDensity(1,(1 + (k - 1 ...
                    )*maxNumExpReps)) = nameBrainRegions(1,k);
            end
        % For the next set of rows, populate with replicate data for each
        % experimental condition
        elseif i > 1 && i <= (1 + height(finalOutputTableCellCounts))
            % Populate all other rows and columns with finalOutputTable
            for j = 1:width(finalOutputTableDensity)
                % If the entry is not equal to 1000, that is as long as the
                % value populating the present cell has been changed from the
                % value to which it was initialised
                if finalOutputTableDensity((i - 1),j) ~= 1000
                    finalOutputTableStrCellCounts(i,j) = num2str( ...
                        finalOutputTableCellCounts((i - 1),j));
                    finalOutputTableStrDensity(i,j) = num2str( ...
                        finalOutputTableDensity((i - 1),j));
                end
            end
        % Populate a blank row after the set of replicate data rows
        elseif i == (height(finalOutputTableDensity) + 1)
            finalOutputTableStrCellCounts(i,:) = "";
            finalOutputTableStrDensity(i,:) = "";
        % For the last set of data rows, and only for the first column per
        % brain region, sequentially populate with mean, SEM, and n
        % summative data for each experimental condition
        elseif i == (height(finalOutputTableDensity) + 2)
            for j = 1:numBrainRegions
                % Running through three rows at a time for the mean, SEM,  
                % and n of each experimental condition
                for n = 1:numCons
                    for m = 1:3
                        % First, populate the mean
                        if m == 1
                            % Extract data subset consisting of all
                            % experimental replicates for the present 
                            % brain region (j) for the present 
                            % experimental condition (n)
                            numExpRepsHere = 0;
                            for p = 1:maxNumExpReps
                                % If the current experimental replicate
                                % data cell being considered is not blank
                                if finalOutputTableCellCounts(n,( ...
                                        (j - 1)*maxNumExpReps + p) ...
                                        ) ~= 1000
                                    % Increment counter
                                    numExpRepsHere = numExpRepsHere + 1;
                                end
                            end
                            nonBlankCellCountsHere = zeros( ...
                                numExpRepsHere,1);
                            nonBlankDensitiesHere = zeros( ...
                                numExpRepsHere,1);
                            for q = 1:maxNumExpReps
                                % Populate with non-zero entries
                                % for mean and SEM calculations
                                if finalOutputTableCellCounts(n,( ...
                                        (j - 1)*maxNumExpReps + q) ...
                                        ) ~= 1000
                                    nonBlankCellCountsHere(q,1 ...
                                        ) =finalOutputTableCellCounts( ...
                                        n,((j - 1)*maxNumExpReps + q));
                                    nonBlankDensitiesHere(q,1 ...
                                        ) = finalOutputTableDensity( ...
                                        n,((j - 1)*maxNumExpReps + q));
                                end
                            end
                            finalOutputTableStrCellCounts((( ...
                                i - 1) + (n - 1)*3 + m + 1),( ...
                                1 + ((j - 1)*maxNumExpReps)) ...
                                ) = num2str(mean(nonBlankCellCountsHere));
                            finalOutputTableStrDensity(( ...
                                (i - 1) + (n - 1)*3 + m + 1),( ...
                                1 + ((j - 1)*maxNumExpReps)) ...
                                ) = num2str(mean(nonBlankDensitiesHere));
                        % Next, populate the SEM. Leave blank if there is
                        % only one value
                        elseif m == 2
                            % If the number of non-blank entries 
                            % calculated at m = 1 is greater than 1
                            if numExpRepsHere > 1
                                finalOutputTableStrCellCounts( ...
                                    ((i - 1) + (n - 1)*3 + m + 1 ...
                                    ),(1 + ((j - 1)*maxNumExpReps) ...
                                    )) = num2str(std( ...
                                    nonBlankCellCountsHere)/sqrt( ...
                                    numExpRepsHere));
                                finalOutputTableStrDensity(( ...
                                    (i - 1) + (n - 1)*3 + m + 1 ...
                                    ),(1 + ((j - 1)*maxNumExpReps) ...
                                    )) = num2str(std( ...
                                    nonBlankDensitiesHere)/sqrt( ...
                                    numExpRepsHere));
                            end
                            % The corresponding cell will be left as a
                            % blank string if there is not more than one
                            % experimental data value
                        elseif m == 3
                            % Finally, populate the n
                            finalOutputTableStrCellCounts(((i - 1 ...
                                ) + (n - 1)*3 + m + 1),(1 + ((j - 1 ...
                                )*maxNumExpReps))) = num2str( ...
                                numExpRepsHere);
                            finalOutputTableStrDensity(((i - 1 ...
                                ) + (n - 1)*3 + m + 1),(1 + (( ...
                                j - 1)*maxNumExpReps)) ...
                                ) = num2str(numExpRepsHere);
                        end
                    end
                end
            end
        end
    end
    
    %% Save output table to the designated Prism directory
    
    % Define the name and location of output cell count and density files
    saveNameCounts = "DREADDs-Processed-" + fluNameHere( ...
        f,1) + "-Counts.csv";
    saveNameDensity = "DREADDs-Processed-" + fluNameHere( ...
        f,1) + "-Density.csv";
    savePathCounts = quintFolderLocation + "\" + saveNameCounts;
    savePathDensity = quintFolderLocation + "\" + saveNameDensity;
    % Save file
    writematrix(finalOutputTableStrCellCounts,savePathCounts)
    writematrix(finalOutputTableStrDensity,savePathDensity)
end

%% Additional re-formatting for multiplexed analyses 

% Ask if the user would like to perform additional re-formatting for 
% stacked/subset analyses
prompt_moreReFormat = "Would you like to conduct additional " + ...
    "re-formatting for compatibility with GraphPAD Prism? " + ...
    "Enter 'Y' if yes, or 'N' if no: ";
moreReFormat = input(prompt_moreReFormat,"s");
% Error-checking for a valid entry
moreReFormat = quint_isentryerror("Y","N",moreReFormat);
if moreReFormat == "Y"
    % Retreive the directory of the QUINT-Prism reformatting file for 
    % navigation
    prompt_templateFolderLocation = "Enter the folder path for the " + ...
        "location of the QUINT-Prism re-formatting Excel sheet: ";
    templateFolderLocation = input(prompt_templateFolderLocation,"s");
    templateFolderLocationChar = convertStringsToChars( ...
        templateFolderLocation);
end
while moreReFormat == "Y"
    % Retreive the name of the input Excel sheet for analysis, which must
    % be in the same folder as that which was specified previously
    prompt_templateExcelFileName = "Enter the name of the " + ...
        "Prism-reformatting template Excel sheet: ";
    templateExcelFileName = input(prompt_templateExcelFileName,"s");
    templateExcelFileNameChar = convertStringsToChars( ...
        templateExcelFileName);
    % Error-checking to ensure the folder and input sheet directory 
    % locations suffice and that the excel data file is not duplicated 
    % in the path specification
    fileNameContainLength = length(templateExcelFileNameChar) - 1;
    fileNameIndexLength = length(templateExcelFileNameChar) + 1;
    % If the folder/file path is long enough to contain the file name
    if length(templateFolderLocationChar) > fileNameContainLength
        % If the file name is included in the user-defined path
        if convertCharsToStrings(templateFolderLocationChar(( ...
                end-fileNameContainLength):end)) == templateExcelFileName
            % Truncate file name and convert data type
            templateFolderLocationChar = templateFolderLocationChar( ...
                1:(end-fileNameIndexLength));
            templateFolderLocation = convertCharsToStrings( ...
                templateFolderLocationChar);
            % File location simply = folder location
            quintExcelFileLocation = templateFolderLocation;
            quintExcelFileLocationChar = templateFolderLocationChar;
        end
    end
    % If the file name was not included in the user-defined path
    if convertCharsToStrings(templateFolderLocationChar(( ...
            end-fileNameContainLength):end)) ~= templateExcelFileName
        % Append fileName
        quintExcelFileLocation = templateFolderLocation + "\" + ...
            "" + templateExcelFileName;
        quintExcelFileLocationChar = convertStringsToChars( ...
            quintExcelFileLocation);
    end
    % Extract data from the input excel sheet after first adding the 
    % folder to the working directory path
    addpath(templateFolderLocation,'-end');
    % Import data
    fprintf("Importing information from ");
    disp(quintExcelFileLocation);
    % Extract excelFileName.txt
    templateReFormatExcelImport = readcell(templateExcelFileNameChar);
    % Extract the number of brain regions of interest here, which is equal
    % to the height of the input table minus three to account for column
    % headings
    numBrainROIsHere = height(templateReFormatExcelImport) - 3;
    % Initialise array to hold the names of the brain regions of interest
    brainROIsHere = strings(numBrainROIsHere,1);
    % Extract brain region names
    for b = 1:numBrainROIsHere
        brainROIsHere(b,1) = convertCharsToStrings( ...
            templateReFormatExcelImport{(3 + b),1});
        % Ensure each brain region exists, iterating through each of the
        % ones analysed by QUINT in the first instance
        doesBRExist = 0;
        for c = 1:numBrainRegions
            if brainROIsHere(b,1) == nameBrainRegions(1,c)
                doesBRExist = 1;
            end
        end
        % Exit the program if the brain region specified does not exist in
        % the QUINT outputs (i.e., if there is an error with user-entry)
        if doesBRExist == 0
            fprintf("Your Prism re-formatting sheet " + ...
                "contains the brain region '" + brainROIsHere( ...
                b,1) + "', which is not recognised in " + ...
                "the present set of QUINT outputs.\n");
            fprintf("Please check your entries and try again.\n");
            return
        end
        % Otherwise, simply repeat through all brain regions and register
        % each in the array brainROIsHere
    end
    % Determine the number of conditions of interest here, starting at
    % column 2 to account for the header in row 1 of the template file
    numCOIsHere = 0;
    for w = 2:width(templateReFormatExcelImport)
        % If the cell is NOT empty
        if ismissing(convertCharsToStrings( ...
                templateReFormatExcelImport{2,w})) == 0
            numCOIsHere = numCOIsHere + 1;
        end
    end
    % Initialise array for conditions of interest here, the first column
    % being for the name and the second column being for the position
    % relative to the QUINT re-formatted dataset (which is essential in
    % then navigating output sheets to extract data)
    nameCOIsHere = strings(numCOIsHere,2);
    % Initialise array incrementer
    nameIndex = 1;
    % Extract the names of the conditions of interest here. Iterate 
    % through each column of the template array and save non-blank entries
    % in the nameCOIsHere array
    for d = 2:width(templateReFormatExcelImport)
        if ismissing(convertCharsToStrings( ...
                templateReFormatExcelImport{2,d})) == 0
            nameCOIsHere(nameIndex,1) = convertCharsToStrings( ...
                templateReFormatExcelImport{2,d});
            % Ensure each experimental condition exists, iterating through 
            % the names of each one registered in the prior QUINT analysis
            doesConExist = 0;
            for h = 1:numCons
                if nameCOIsHere(nameIndex,1) == nameCons(h,1)
                    nameCOIsHere(nameIndex,2) = num2str(h);
                    doesConExist = 1;
                end
            end
            % Exit the program if the condition specified does not exist 
            % in QUINT outputs (i.e., an error with user-entry)
            if doesConExist == 0
                fprintf("Your Prism re-formatting sheet " + ...
                    "contains the condition '" + nameCOIsHere( ...
                    nameIndex,1) + "', which is not recognised " + ...
                    "in the present set of QUINT outputs.\n");
                fprintf("Please check your entries and try again.\n");
                return
            end
            nameIndex = nameIndex + 1;
        end
        % Otherwise, simply repeat through conditions and register each in 
        % the array nameCOIsHere
    end
    % Provide a nickname for the present brain region of interest
    % combination
    fprintf("Enter a short name/ID for the " + ...
        "brain region combination represented " + ...
        "in this file:\n");
    fprintf(templateExcelFileName + ": ");
    prompt_brainROIComboID = "This will be included in the name " + ...
        "of the output file in the format " + ...
        "(your-entry-here)_fluorophore-name_Density.csv: ";
    brainROIComboID = input(prompt_brainROIComboID,"s");
    % For each fluorophore
    for flu = 1:numFluoroCombos
        % Extract the output table
        quintOutputforFluorHereName = "DREADDs-Processed-" + ...
            "" + fluNameHere(flu,1) + "-Density.csv";
        quintOutputforFluorHereNameChar = convertStringsToChars( ...
            quintOutputforFluorHereName);
        % Extract QUINT output data for the fluorophore of interest
        quintOutputFluorTableHere = readcell( ...
            quintOutputforFluorHereNameChar);
        % Initialise array to hold extracted data
        dataExtract = zeros(numBrainROIsHere,(numCOIsHere*3));
        % For each brain region of interest
        for b = 1:numBrainROIsHere
            % Scan through column headers until the brain region of
            % interest is located
            for h = 1:width(quintOutputFluorTableHere)
                % Locating the target column
                if brainROIsHere(b,1) == convertCharsToStrings( ...
                        quintOutputFluorTableHere{1,h})
                    % Parse through the corresponding column to summative
                    % Mean/SEM/n information, specifically to the
                    % experimental conditions in the template sheet,
                    % starting at row 2 + numCons to account for column
                    % headers and the first set of replicate data rows per
                    % experimental condition. Then add 3*(index of the
                    % specified experimental condition) and increment by
                    % one respectively to find the Mean, SEM, and n for 
                    % the present condition (iterating through each
                    % condition)
                    for t = 1:numCOIsHere
                        % Log the mean
                        dataExtract(b,((t - 1)*3 + 1) ...
                            ) = quintOutputFluorTableHere{(2 + ...
                            numCons + (str2double(nameCOIsHere(t,2) ...
                            ) - 1)*3),h};
                        % Log the SEM, which is one row beneath the mean
                        % for the current brain region of interest and
                        % experimental condition
                        dataExtract(b,((t - 1)*3 + 2) ...
                            ) = quintOutputFluorTableHere{(2 + numCons ...
                            + (str2double(nameCOIsHere(t,2)) - 1 ...
                            )*3 + 1),h};
                        % Log the n, which is one row beneath the SEM for
                        % the current brain region of interest and
                        % experimental condition
                        dataExtract(b,((t - 1)*3 + 3) ...
                            ) = quintOutputFluorTableHere{( ...
                            2 + numCons + (str2double(nameCOIsHere( ...
                            t,2)) - 1)*3 + 2),h};
                        % Repeating for each experimental condition 
                    end
                end
            end
            % Repeating for each brain region specified
        end
        % Adding row and column headers to dataExtract
        dataExtractRC = strings((height(dataExtract) + 1),( ...
            width(dataExtract) + 1));
        % Populating the first column with the regions of interest
        dataExtractRC(2:end,1) = brainROIsHere(:,1);
        % Populating the first row with the conditions of interest in the
        % appropriate columns
        for cons = 1:numCOIsHere
            dataExtractRC(1,(2 + (cons - 1)*3)) = nameCOIsHere(cons,1);
        end
        % Populating all other data rows and columns with dataExtract
        dataExtractRC(2:end,2:end) = dataExtract(:,:);
        % Name the folder where all re-formatted data will be saved
        quintImportFolder = "QUINT-Output_Prism_Input";
        %addpath(quintFolderLocation,'-end');
        prismDataFolderLocation = templateFolderLocation + "\" + ...
            "" + quintImportFolder;
        % Make a new folder where outputs will be saved, if the folder  
        % does not already exist, and add it to the working directory
        if isfolder(quintImportFolder) == 0
           mkdir QUINT-Output_Prism_Input 
           % Add this path to the working directory
           addpath(quintImportFolder,'-end');
        end
        % Name output table
        saveNamePrismReFormatDensity = brainROIComboID + "_" + ...
            "" + fluNameHere(flu,1) + "_Density.csv";
        savePathPrismReFormatDensity = prismDataFolderLocation + "\" + ...
            "" + saveNamePrismReFormatDensity;
        % Save file
        writematrix(dataExtractRC,savePathPrismReFormatDensity)
    % Repeating for each fluorophore combination
    end   
    prompt_againReFormat = "Would you like to conduct additional " + ...
        "re-formatting for another data combination? " + ...
        "Enter 'Y' if yes, or 'N' if no: ";
    moreReFormat = input(prompt_againReFormat,"s");
    % Error-checking for a valid entry
    moreReFormat = quint_isentryerror("Y","N",moreReFormat);
    % Repeating for other brain region combinations, if specified
end
fprintf("\n");
fprintf("Data has been processed and saved.\n");
fprintf("Program complete.\n");

%% End-of-script functions

function checkedvar = quint_isentryerror(validentry1,validentry2,checkvar)

% Ensure entries are of data type 'string' for comparison. Inputs 
% may not be assigned as strings in the parent script's function call
validentry1 = string(validentry1);
validentry2 = string(validentry2);
checkvar = string(checkvar);
% Re-prompts for a valid response while the user-response does not match 
% a valid entry
while checkvar ~= validentry1 && checkvar ~= validentry2
    prompt_invalidentry = "Invalid entry. Please enter '" + ...
        "" + validentry1 + " or '" + validentry2 + "': ";
    checkvar = input(prompt_invalidentry, "s");
end
checkedvar = string(checkvar);
end