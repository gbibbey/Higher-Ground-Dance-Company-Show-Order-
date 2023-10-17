function HGDC()

%% Read the Inputs from the user
prompt = {'Enter Input Excel Name:','Enter Output Excel Name:', 'Enter Column Number for First Song:', 'Enter Column Number for Last Song:', 'Enter Column Number for Song Before Intermission:', 'Enter Column Number for Song After Intermission:', 'Enter Desired Number of Different Ways to Sort:'};
% dlgtitle = 'Inputs';
% dims = [1 55];
% definput = {'20','hsv'};
% answer = inputdlg(prompt,dlgtitle,dims,definput);
answers = inputdlg(prompt);

% tableName = 'Show Order Spring 22.xlsx';
% outputName = 'Sorted_Dances.xlsx';
tableName = string(answers(1));
outputName = string(answers(2));
firstCol = str2double(answers(3));
endCol = str2double(answers(4));
bInterCol = str2double(answers(5));
aInterCol = str2double(answers(6));
numOutputs = str2double(answers(7));


%% Read the Excel Table 
t = readtable(tableName, 'VariableNamingRule','preserve');
noNames = table2cell(t);
unsorted = [t.Properties.VariableNames; noNames];

%% Delete Output Excel File
delete(outputName);

%% Get First, End, and Intermissions Dances
cols = size(unsorted, 2); % Number of Columns

% Parse the Dances 
first = unsorted(:,firstCol); % First Dance
firstE = first(~cellfun(@isempty, first));
binter = unsorted(:,bInterCol); % Before Intermission Dance
binterE = binter(~cellfun(@isempty, binter));
ainter = unsorted(:,aInterCol); % After Intermission Dance
ainterE = ainter(~cellfun(@isempty, ainter));
endin = unsorted(:,endCol); % End Dance
endinE = endin(~cellfun(@isempty, endin));
 
%% Get x Different Sorted Dances
for m = 1:numOutputs
    %% Go through all the permutations of dances until it fits criteria
    counter = 0;
    while 1 %Keep going until there is a graph that functyions
        P = randperm(cols);
        random = unsorted(:, P); %create random order of dances 
        sorted = cell(size(unsorted));
        sorted(:,1) = first;
        sorted(:,cols/2)= binter;
        sorted(:,cols/2 +1) = ainter;
        sorted(:, cols) = endin;
        for index = 2:(cols - 1) %Create a sorted version of dance
            %Get the prev Dance
            prevDance = sorted(:, index - 1);
            prevDanceE = prevDance(~cellfun(@isempty, prevDance));
            if (isempty(prevDanceE))
                break
            end
            checkAhead = 0;
            if (index == cols/2 || index == (cols/2 + 1))
                continue
            end
            if (index + 1 == cols/2 || index + 1 == cols)
                checkAhead = 1;
                nextDance = sorted(:,index + 1);
                nextDanceE = nextDance(~cellfun(@isempty, nextDance));
            end
            for i = 1:cols %go through all columns in the random matrix to see if any dance can be inserted
                danceToInst = random(:,i);
                danceToInstE = danceToInst(~cellfun(@isempty, danceToInst));
                danceAlreadySorted = 0;
                for j = 1:cols %go through all columns to see if already there
                    danceName = sorted(:, j);
                    danceNameE = danceName(~cellfun(@isempty, danceName));
                    if(isempty(danceNameE))
                        continue
                    end
                    if(strcmp(danceToInstE(1),danceNameE(1))) %If the title of the dance is already there 
                        danceAlreadySorted = 1;
                        break
                    end
                end
                 
                %If Prev and Current Dance have same Type or if Dance already
                %sorted go to next dance    
                if(danceAlreadySorted == 1 || strcmp(danceToInstE(2),prevDanceE(2))) %If the Type of the dance is already there go to next dance
                    continue
                end
                if(checkAhead == 1 && strcmp(danceToInstE(2), nextDanceE(2))) %Check upcoming dances
                    continue
                end
    
                sameDancer = 0;
                for j = 3:length(prevDanceE) %Check all dancers
                    for k = 3:length(danceToInstE)
                        if(strcmp(prevDanceE(j), danceToInstE(k)))
                            sameDancer = 1;
                            break
                        end
                    end
                    if (sameDancer == 1)
                        break
                    end
                end
                 
                if (checkAhead == 1 && sameDancer == 0)
                    for j = 3:length(nextDanceE) %Check all dancers Ahead
                        for k = 3:length(danceToInstE)
                            if(strcmp(nextDanceE(j), danceToInstE(k)))
                                sameDancer = 1;
                                break
                            end
                        end
                        if (sameDancer == 1)
                            break
                        end
                    end
                end
                 
                if (sameDancer == 1)
                    continue
                else
                    sorted(:, index) = danceToInst;
                    break
                end
            end
        end
        newRand = 0;
        for sortIndex = 1:cols
            dance = sorted(:, sortIndex);
            danceE = dance(~cellfun(@isempty, dance));
            if(isempty(danceE))
                counter = counter + 1;
                newRand = 1;
                break
            end
        end
        if (newRand == 0)
            break
        end
    end
    
    sortedT = cell2table(sorted(2:end, :));
    sortedT.Properties.VariableNames = sorted(1,:);
    writetable(sortedT,outputName,'Sheet',m);
end

msgbox("Sorting Completed", "Success");