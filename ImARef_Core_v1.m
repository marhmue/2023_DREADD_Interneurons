%{

QUINT Data Processing
Molnar Lab 2023
Marissa Mueller

quint_postprocessing_subdivisions.m

%}

%{

Overview: 

This code imports QUINT pipeline output data files, then re-formats for
data visualisation and analysis in GraphPAD PRISM. It returns both cell
count and density information for all inputted  fluorophore combinations 
of interest

Code Requirements: 

QUINT output directory location and user-inputted processing 
specifications (see Input). If Prism re-formatting is selected, the user
must also have a directory in the parent folder which contains template
Excel sheets and a sub-directory named "QUINT-Output_Prism_Input" which
will house final re-formatted outputs

Inputs:

 - quintFolderLocation (parent file location for QUINT outputs)
 - numCons (the number of conditions/genotypes being assessed)
 - nameCons (name each condition/genotype)
 - numExpReps (number of experimental replicates/animals per condition)
 - nameExpReps (name each experimental replicate/animal)
 - numTechReps (the number of technical replicates/sections per 
   experimental replicate/animal)
 - numBrainRegions (the number of brain regions assessed by QUINT outputs)

Outputs:

 - Processed QUINT output data (Excel sheets for Prism import)

Export and Application:

Outputs are saved in the parent folder as QUINT-Processed.txt. This file 
(.csv) can be exported to a graphing program of choice for further
visualisation and data analysis. Columns are grouped by brain region, with
condition/genotype status being discriminated by row. Experimental
replicates (animals) exist as sub-columns with technical replicates 
(repeat slices) collapsed into single averaged measurements. Summary data
are provided underneath in the order of genotype appearance (mean, SEM, N)
for graphing in Prism

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
    % Retreive the name of the instructional input Excel sheet for 
    % analysis (i.e., which contains information regarding experimental 
    % design, experimental replicates, and technical replicates
    prompt_quintExcelFileName = "Enter the name of the instructional input Excel sheet being used for analysis: ";
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
            quintFolderLocationChar = quintFolderLocationChar(1:(end-fileNameIndexLength));
            quintFolderLocation = convertCharsToStrings(quintFolderLocationChar);
            % quintExcelFileLocation simply = quintFolderLocation
            quintExcelFileLocation = quintFolderLocation;
            quintExcelFileLocationChar = quintFolderLocationChar;
        end
    end
    % If excelFileName was not included in the user-defined path
    if convertCharsToStrings(quintFolderLocationChar((end-fileNameContainLength):end)) ~= excelFileName
        % The folder string is already defined as quintFolderLocation
        % Append fileName
        quintExcelFileLocation = quintFolderLocation + "\" + excelFileName;
        quintExcelFileLocationChar = convertStringsToChars(quintExcelFileLocation);
    end
    % Extract data from the input excel sheet
    % Add the skrtkdFileLocation folder to the working directory path
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
        nameConHere = convertCharsToStrings(quintExcelImport{(11 + (i - 1)*maxNumExpReps),1});
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
            nameExpReps = convertCharsToStrings(quintExcelImport{(11 + (i - 1)*maxNumExpReps + (j - 1)),3});
            nameExpRepsRow = conRow + (j - 1);
            dataScaffoldStr(nameExpRepsRow,2) = nameExpReps;
            % Extract the number of technical replicates for this 
            % condition and experimental replicate
            numTechReps = quintExcelImport{(11 + (i - 1)*maxNumExpReps + (j - 1)),4};
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
        fluNameHere(g,1) = convertCharsToStrings(quintExcelImport{(11 + (g - 1)),6});
    end
    % Extract the the area (microns, um^2) corresponding to one pixel
    squareMicronsPerPixel = quintExcelImport{8,6};
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
    fprintf("Microns per pixel = " + squareMicronsPerPixel + " um^2/px");
    fprintf("\n");
    % Ask whether these are correct (i.e., whether the program should proceed
    % or whether inputs need to be re-entered)
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
    % Initialise an array for final output data. Columns will represent data
    % from experimental replicates for each brain region. Rows will represent
    % information retrieved for each experimental condition/genotype
    numColumnsFinal = numBrainRegions*maxNumExpReps;
    numRowsFinal = numCons;
    finalOutputTableCellCounts = ones(numRowsFinal,numColumnsFinal).*1000;
    finalOutputTableDensity = ones(numRowsFinal,numColumnsFinal).*1000;
    finalOutputTableAreas = ones(numRowsFinal,numColumnsFinal).*1000;
    % Iterate through directories to extract QUINT outputs
    % For each experimental condition/genotype
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
            intFPath = quintFolderLocation + "\" + expRepHere + "\8_Nutil_Output\" + fluNameHere(f,1) + "\Reports\";
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
            % Specify the standardised file name according to QUINT
            % formatting
            fileName = fileFolder + "_All";
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
                    nameBrainRegions(1,(q - 1)) = convertCharsToStrings(quintImport{q,1});
                end
            end
            % For each brain region, starting at row p = 2 to account for
            % column headers
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
                % Calculate density, being sure that pixel values are
                % accounted for here irrespective of whether they are
                % defined/accounted for in the Nutil quantifier
                regionAreaHere = regionPixelsHere*squareMicronsPerPixel;
                densityHere = numObjectsHere/regionAreaHere;
                % Calculate object average area
                objectPixelsImport = quintImport(p,6);
                % Convert to numerical format
                objectPixelsHere = cell2mat(objectPixelsImport);
                % Calculate average area per cell body (object)
                objectAreasHere = objectPixelsHere*squareMicronsPerPixel/numObjectsHere;
                % Populate output table
                finalOutputTableCellCounts(i,(j + (p - 2)*maxNumExpReps)) = numObjectsHere;
                finalOutputTableDensity(i,(j + (p - 2)*maxNumExpReps)) = densityHere;
                finalOutputTableAreas(i,(j + (p - 2)*maxNumExpReps)) = objectAreasHere;
            end
        end
    end
    
    % Re-format output table (featuring easy import to Prism)
    
    % For the number of rows, include (in the order of appearance in the
    % equaltion) a column header, a data row for each experimental
    % condition, a blank 'spacer' row, and three rows (in the first column 
    % of each brain region) per experimental condition for the mean, SEM, 
    % and n of technical replicates. This last set of rows is constructed
    % for compatibility with GraphPad prism and ease of data transfer
    finalOutputTableStrCellCounts = strings((1 + height(finalOutputTableCellCounts) + 1 + numCons*3),width(finalOutputTableCellCounts));
    finalOutputTableStrDensity = strings((1 + height(finalOutputTableCellCounts) + 1 + numCons*3),width(finalOutputTableDensity));
    finalOutputTableStrCellAreas = strings((1 + height(finalOutputTableCellCounts) + 1 + numCons*3),width(finalOutputTableAreas));
    for i = 1:(height(finalOutputTableDensity) + 2)
        % For the first row, simply populate column headers with the name
        % of each brain region
        if i == 1 
            for k = 1:numBrainRegions
                finalOutputTableStrCellCounts(1,(1 + (k - 1)*maxNumExpReps)) = nameBrainRegions(1,k);
                finalOutputTableStrDensity(1,(1 + (k - 1)*maxNumExpReps)) = nameBrainRegions(1,k);
                finalOutputTableStrCellAreas(1,(1 + (k - 1)*maxNumExpReps)) = nameBrainRegions(1,k);
            end
        % For the next set of rows, populate with replicate data for each
        % experimental condition
        elseif i > 1 && i <= (1 + height(finalOutputTableCellCounts))
            % Populate all other rows and columns with finalOutputTable
            for j = 1:width(finalOutputTableDensity)
                finalOutputTableStrCellCounts(i,j) = num2str(finalOutputTableCellCounts((i - 1),j));
                finalOutputTableStrDensity(i,j) = num2str(finalOutputTableDensity((i - 1),j));
                finalOutputTableStrCellAreas(i,j) = num2str(finalOutputTableAreas((i - 1),j));
            end
        % Populate a blank row after the set of replicate data rows
        elseif i == (height(finalOutputTableDensity) + 1)
            finalOutputTableStrCellCounts(i,:) = "";
            finalOutputTableStrDensity(i,:) = "";
            finalOutputTableStrCellAreas(i,:) = "";
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
                                if finalOutputTableCellCounts(n,((j - 1)*maxNumExpReps + p)) ~= 1000
                                    % Increment counter
                                    numExpRepsHere = numExpRepsHere + 1;
                                end
                            end
                            nonBlankCellCountsHere = zeros(numExpRepsHere,1);
                            nonBlankDensitiesHere = zeros(numExpRepsHere,1);
                            nonBlankAreasHere = zeros(numExpRepsHere,1);
                            for q = 1:maxNumExpReps
                                % Populate with non-zero entries
                                % for mean and SEM calculations
                                if finalOutputTableCellCounts(n,((j - 1)*maxNumExpReps + q)) ~= 1000
                                    nonBlankCellCountsHere(q,1) = finalOutputTableCellCounts(n,((j - 1)*maxNumExpReps + q));
                                    nonBlankDensitiesHere(q,1) = finalOutputTableDensity(n,((j - 1)*maxNumExpReps + q));
                                    nonBlankAreasHere(q,1) = finalOutputTableAreas(n,((j - 1)*maxNumExpReps + q));
                                end
                            end
                            finalOutputTableStrCellCounts(((i - 1) + (n - 1)*3 + m + 1),(1 + ((j - 1)*maxNumExpReps))) = num2str(mean(nonBlankCellCountsHere));
                            finalOutputTableStrDensity(((i - 1) + (n - 1)*3 + m + 1),(1 + ((j - 1)*maxNumExpReps))) = num2str(mean(nonBlankDensitiesHere));
                            finalOutputTableStrCellAreas(((i - 1) + (n - 1)*3 + m + 1),(1 + ((j - 1)*maxNumExpReps))) = num2str(mean(nonBlankAreasHere));
                        % Next, populate the SEM. Leave blank if there is
                        % only one value
                        elseif m == 2
                            % If the number of non-blank entries 
                            % calculated at m = 1 is greater than 1
                            if numExpRepsHere > 1
                                finalOutputTableStrCellCounts(((i - 1) + (n - 1)*3 + m + 1),(1 + ((j - 1)*maxNumExpReps))) = num2str(std(nonBlankCellCountsHere)/sqrt(numExpRepsHere));
                                finalOutputTableStrDensity(((i - 1) + (n - 1)*3 + m + 1),(1 + ((j - 1)*maxNumExpReps))) = num2str(std(nonBlankDensitiesHere)/sqrt(numExpRepsHere));
                                finalOutputTableStrCellAreas(((i - 1) + (n - 1)*3 + m + 1),(1 + ((j - 1)*maxNumExpReps))) = num2str(std(nonBlankAreasHere)/sqrt(numExpRepsHere));
                            end
                            % The corresponding cell will be left as a
                            % blank string if there is not more than one
                            % experimental data value
                        % Finally, populate the n
                        elseif m == 3
                            finalOutputTableStrCellCounts(((i - 1) + (n - 1)*3 + m + 1),(1 + ((j - 1)*maxNumExpReps))) = num2str(numExpRepsHere);
                            finalOutputTableStrDensity(((i - 1) + (n - 1)*3 + m + 1),(1 + ((j - 1)*maxNumExpReps))) = num2str(numExpRepsHere);
                            finalOutputTableStrCellAreas(((i - 1) + (n - 1)*3 + m + 1),(1 + ((j - 1)*maxNumExpReps))) = num2str(numExpRepsHere);
                        end
                    end
                end
            end
        end
    end
    
    %% Save output table to the designated Prism directory
    
    % Define the name and location of output cell count and density files
    saveNameCounts = "Processed-" + fluNameHere(f,1) + "-Counts.csv";
    saveNameDensity = "Processed-" + fluNameHere(f,1) + "-Density.csv";
    saveNameAreas = "Processed-" + fluNameHere(f,1) + "-Areas.csv";
    savePathCounts = quintFolderLocation + "\" + saveNameCounts;
    savePathDensity = quintFolderLocation + "\" + saveNameDensity;
    savePathAreas = quintFolderLocation + "\" + saveNameAreas;
    % Save file
    writematrix(finalOutputTableStrCellCounts,savePathCounts)
    writematrix(finalOutputTableStrDensity,savePathDensity)
    writematrix(finalOutputTableStrCellAreas,savePathAreas)
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
        "location of QUINT-Prism re-formatting Excel sheet(s): ";
    templateFolderLocation = input(prompt_templateFolderLocation,"s");
    templateFolderLocationChar = convertStringsToChars(templateFolderLocation);
    % Extract the names of all excel sheets
    excelFileNames = dir([templateFolderLocationChar, '\*.xlsx']);
    numExcelFiles = length(excelFileNames);
end
for e = 1:numExcelFiles
    % Retreive the name of the input Excel sheet for analysis. These
    % are blank templates which will be populated for direct import
    % thereafter to Prism
    templateExcelFileNameChar = excelFileNames(e).name;
    templateExcelFileName = convertCharsToStrings(templateExcelFileNameChar);
    % Extract data from the input excel sheet
    % Add the folder to the working directory path
    addpath(templateFolderLocation,'-end');
    % Import QUINT-Analysis files
    fprintf("Importing information from " + templateExcelFileName);
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
        if ismissing(convertCharsToStrings(templateReFormatExcelImport{2,w})) == 0
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
        if ismissing(convertCharsToStrings(templateReFormatExcelImport{2,d})) == 0
            nameCOIsHere(nameIndex,1) = convertCharsToStrings(templateReFormatExcelImport{2,d});
            % Ensure each experimental condition exists, iterating through the
            % names of each experimental condition registered in the QUINT
            % analysis previously conducted
            doesConExist = 0;
            for h = 1:numCons
                if nameCOIsHere(nameIndex,1) == nameCons(h,1)
                    nameCOIsHere(nameIndex,2) = num2str(h);
                    doesConExist = 1;
                end
            end
            % Exit the program if the bcondition specified does not exist in
            % the QUINT outputs (i.e., if there is an error with user-entry)
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
    fprintf("... ")
    % For each fluorophore
    for flu = 1:numFluoroCombos
        % Extract the output table, first for density then for cell counts
        % then for areas
        quintOutputforFluorHereName = "Processed-" + fluNameHere(flu,1) + "-Density.csv";
        quintOutputforFluorHereNameChar = convertStringsToChars(quintOutputforFluorHereName);
        quintOutputforFluorHereNameCounts = "Processed-" + fluNameHere(flu,1) + "-Counts.csv";
        quintOutputforFluorHereNameCharCounts = convertStringsToChars(quintOutputforFluorHereNameCounts);
        quintOutputforFluorHereNameAreas = "Processed-" + fluNameHere(flu,1) + "-Areas.csv";
        quintOutputforFluorHereNameCharAreas = convertStringsToChars(quintOutputforFluorHereNameAreas);
        % Extract QUINT output data for the present fluorophore of interest
        quintOutputFluorTableHere = readcell(quintOutputforFluorHereNameChar);
        quintOutputFluorTableHereCounts = readcell(quintOutputforFluorHereNameCharCounts);
        quintOutputFluorTableHereAreas = readcell(quintOutputforFluorHereNameCharAreas);
        % Initialise array to hold extracted data
        dataExtract = zeros(numBrainROIsHere,(numCOIsHere*3));
        dataExtractCounts = zeros(numBrainROIsHere,(numCOIsHere*3));
        dataExtractAreas = zeros(numBrainROIsHere,(numCOIsHere*3));
        % For each brain region of interest
        for b = 1:numBrainROIsHere
            % Scan through column headers until the brain region of
            % interest is located
            for h = 1:width(quintOutputFluorTableHere)
                % Locate the target column, which applies to both denities
                % and counts as they have the same format
                if brainROIsHere(b,1) == convertCharsToStrings(quintOutputFluorTableHere{1,h})
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
                        % Changing the 2 to a 1 due to formatting
                        % differences - change back to 2 if using data 
                        % from previous verions of this code
                        rowIndicator = 2;
                        % Log the mean
                        dataExtract(b,((t - 1)*3 + 1)) = quintOutputFluorTableHere{(rowIndicator + numCons + (str2double(nameCOIsHere(t,2)) - 1)*3),h};
                        dataExtractCounts(b,((t - 1)*3 + 1)) = quintOutputFluorTableHereCounts{(rowIndicator + numCons + (str2double(nameCOIsHere(t,2)) - 1)*3),h};
                        dataExtractAreas(b,((t - 1)*3 + 1)) = quintOutputFluorTableHereAreas{(rowIndicator + numCons + (str2double(nameCOIsHere(t,2)) - 1)*3),h};
                        % Log the SEM, which is one row beneath the mean
                        % for the current brain region of interest and
                        % experimental condition
                        dataExtract(b,((t - 1)*3 + 2)) = quintOutputFluorTableHere{(rowIndicator + numCons + (str2double(nameCOIsHere(t,2)) - 1)*3 + 1),h};
                        dataExtractCounts(b,((t - 1)*3 + 2)) = quintOutputFluorTableHereCounts{(rowIndicator + numCons + (str2double(nameCOIsHere(t,2)) - 1)*3 + 1),h};
                        dataExtractAreas(b,((t - 1)*3 + 2)) = quintOutputFluorTableHereAreas{(rowIndicator + numCons + (str2double(nameCOIsHere(t,2)) - 1)*3 + 1),h};
                        % Log the n, which is one row beneath the SEM for
                        % the current brain region of interest and
                        % experimental condition
                        dataExtract(b,((t - 1)*3 + 3)) = quintOutputFluorTableHere{(rowIndicator + numCons + (str2double(nameCOIsHere(t,2)) - 1)*3 + 2),h};
                        dataExtractCounts(b,((t - 1)*3 + 3)) = quintOutputFluorTableHereCounts{(rowIndicator + numCons + (str2double(nameCOIsHere(t,2)) - 1)*3 + 2),h};
                        dataExtractAreas(b,((t - 1)*3 + 3)) = quintOutputFluorTableHereAreas{(rowIndicator + numCons + (str2double(nameCOIsHere(t,2)) - 1)*3 + 2),h};
                        % Repeating for each experimental condition specified
                    end
                end
            end
            % Repeating for each brain region specified
        end
        % Adding row and column headers to dataExtract
        dataExtractRC = strings((height(dataExtract) + 1),(width(dataExtract) + 1));
        dataExtractRCCounts = strings((height(dataExtractCounts) + 1),(width(dataExtractCounts) + 1));
        dataExtractRCAreas = strings((height(dataExtractCounts) + 1),(width(dataExtractCounts) + 1));
        % Populating the first column with the regions of interest
        dataExtractRC(2:end,1) = brainROIsHere(:,1);
        dataExtractRCCounts(2:end,1) = brainROIsHere(:,1);
        dataExtractRCAreas(2:end,1) = brainROIsHere(:,1);
        % Populating the first row with the conditions of interest in the
        % appropriate columns
        for cons = 1:numCOIsHere
            dataExtractRC(1,(2 + (cons - 1)*3)) = nameCOIsHere(cons,1);
            dataExtractRCCounts(1,(2 + (cons - 1)*3)) = nameCOIsHere(cons,1);
            dataExtractRCAreas(1,(2 + (cons - 1)*3)) = nameCOIsHere(cons,1);
        end
        % Populating all other data rows and columns with dataExtract
        dataExtractRC(2:end,2:end) = dataExtract(:,:);
        dataExtractRCCounts(2:end,2:end) = dataExtractCounts(:,:);
        dataExtractRCAreas(2:end,2:end) = dataExtractAreas(:,:);
        % Name the folder where all re-formatted data will be saved
        quintImportFolder = "QUINT-Output_Prism_Input";
        %addpath(quintFolderLocation,'-end');
        prismDataFolderLocation = templateFolderLocation + "\" + quintImportFolder;
        % Make a new folder where outputs will be saved, if the folder does 
        % not already exist, and add it to the working directory
        if isfolder(quintImportFolder) == 0
           mkdir QUINT-Output_Prism_Input_2025 
           % Add this path to the working directory
           addpath(quintImportFolder,'-end');
        end
        % Name output tables and paths
        saveNamePrismReFormatDensity = templateExcelFileName + "_" + fluNameHere(flu,1) + "_Density.csv";
        savePathPrismReFormatDensity = prismDataFolderLocation + "\" + saveNamePrismReFormatDensity;
        saveNamePrismReFormatCounts = templateExcelFileName + "_" + fluNameHere(flu,1) + "_Counts.csv";
        savePathPrismReFormatCounts = prismDataFolderLocation + "\" + saveNamePrismReFormatCounts;
        saveNamePrismReFormatAreas = templateExcelFileName + "_" + fluNameHere(flu,1) + "_Areas.csv";
        savePathPrismReFormatAreas = prismDataFolderLocation + "\" + saveNamePrismReFormatAreas;
        % Save files
        writematrix(dataExtractRC,savePathPrismReFormatDensity)
        writematrix(dataExtractRCCounts,savePathPrismReFormatCounts)
        writematrix(dataExtractRCAreas,savePathPrismReFormatAreas)
    % Repeating for each fluorophore combination
    end
    fprintf("Complete \n");
    % Repeating for all brain region combinations defined by Excel sheets
end
fprintf("\n");
fprintf("Data has been processed and saved.\n");
fprintf("Program complete.\n");
% End of script 

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