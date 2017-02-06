clear;
close all;

% Set start and end day
startDate = datetime(2015, 1, 1);
endDate = datetime(2015, 12, 31);
% Determine how many days will be simulated
dayCount = length(startDate:endDate);

% Make 100 dwelling datasets
for profile_count = 0:99
    
    % Connect to excel application
    try
        ExcelApp = actxGetRunningServer('Excel.Application');
    catch
        disp('Please make sure Excel is opened');
        return
    end
    % Makes window visible
    ExcelApp.Visible = 1;
    % Opens workbook
    Workbook = ExcelApp.Workbooks.Open(fullfile(pwd, '\modified.xlsm'));

    % Set random number of residents per house
    resident_count = randperm(5, 1);
    ExcelApp.Range('main!$K$4').Value = resident_count;
    
    % Allocates appliances to this dwelling only once
    ExcelApp.Run('Sheet1.btnAllocateAppliances_Click');
        
    annual_load_profile = zeros(1440 * 365, 1);
    annual_irradiance_profile = zeros(1440 * 365, 1);
    
    % Start or resets the timer
    tic
    
    % Steps through the specified range (here year 2015)
    for d = startDate:endDate
        
        if isweekend(d)
            % Set weekend
            ExcelApp.Range('main!$K$5').Value = 'we';
        else
            % Set weekday
            ExcelApp.Range('main!$K$5').Value = 'wd';
        end
        
        % Set the current month
        ExcelApp.Range('main!$K$6').Value = month(d);
        
        % Runs the occupancy simulation
        ExcelApp.Run('Sheet1.btnRunOccupancy_Click');
        % Runs the electricity demand model
        ExcelApp.Run('Sheet1.btnRunApplianceModel_Click');
        
        % Get the day's number in the year
        dayIndex = 1440*(day(d, 'dayofyear')-1)+1;
        
        % Extracting load data (kW)
        DwellingLoad = ExcelApp.Range('appliance_sim_data!$D$12:$D$1451');
        DwellingLoadProfile = cell2mat(DwellingLoad.Value)./1000;
        
        annual_load_profile(dayIndex:dayIndex+1439) = DwellingLoadProfile;
        
        % Extract irradiance data (kW per m^2)
        Irradiance = ExcelApp.Range('irradiance!$C$12:$C$1451');
        IrradianceProfile = cell2mat(Irradiance.Value)./1000;
        
        annual_irradiance_profile(dayIndex:dayIndex+1439) = IrradianceProfile;
        
        % Delete old progress
        if exist('print_str', 'var')
            for i = 1:length(print_str)
                fprintf('\b');
            end
        end
        
        % Give new progress update
        elapsed = toc;
        progress = day(d, 'dayofyear') / dayCount;
        remaining = (1-progress)*elapsed/progress / 60;
        print_str = sprintf('Finished %s [%6.2f%%] : %.2f minutes',...
            datestr(d), progress*100, remaining);
        fprintf('%s', print_str);
    end
    
    % Save data for dwelling
    save(sprintf('dwelling_%03d.mat', profile_count), ...
        'annual_*', 'resident_count');
    
    % Tidy up and close Excel
    Workbook.Saved = 1;
    ExcelApp.Quit;
    ExcelApp.release;
    
    fprintf(' -> Finished dwelling %03d\n', profile_count);
    clear print_str
    
end
