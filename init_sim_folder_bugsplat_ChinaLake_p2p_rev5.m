clear;
clc;
close all force;
close all;
app=NaN(1);  %%%%%%%%%This is to allow for Matlab Application integration.
format shortG
top_start_clock=clock;
folder1='C:\Users\nlasorte\OneDrive - National Telecommunications and Information Administration\MATLAB2024\7GHz P2P Single Entry';
cd(folder1)
addpath(folder1)
addpath('C:\Users\nlasorte\OneDrive - National Telecommunications and Information Administration\MATLAB2024\Basic_Functions') %%%%%%%%This is another Github repo
addpath('C:\Users\nlasorte\OneDrive - National Telecommunications and Information Administration\MATLAB2024\General_Movelist')  %%%%%%%%This is another Github repo
addpath('C:\Users\nlasorte\OneDrive - National Telecommunications and Information Administration\MATLAB2024\General_Terrestrial_Pathloss')  %%%%%%%%This is another Github repo
addpath('C:\Users\nlasorte\OneDrive - National Telecommunications and Information Administration\MATLAB2024\Generic_Bugsplat') %%%%%%%%This is another Github repo
addpath('C:\Users\nlasorte\OneDrive - National Telecommunications and Information Administration\MATLAB2024\Census_Functions')
addpath('C:\Local Matlab Data\Local MAT Data') %%%%%%%One Drive Error with mat files
pause(0.1)

'Pull the new spreadsheet for China Lake'

'P2P FDR Bugsplats'
'1.Pull the database'
'2.Calculate FDR for each channel'

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%NSF Data
tf_repull_p2p_data=0%1%0%1%0%1
data_num=4%3%2%1

cell_data_filename=strcat('ChinaLake_P2P_Adjacent_cell_sim_data',num2str(data_num),'.mat');
[var_exist]=persistent_var_exist_with_corruption(app,cell_data_filename);
if tf_repull_p2p_data==1
    var_exist=0;
end
if var_exist==2
    load(cell_data_filename,'cell_sim_data')
else

    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Data Header
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    cell_data_header=cell(1,42);
    cell_data_header{1}='data_label1';
    cell_data_header{2}='latitude';
    cell_data_header{3}='longitude';
    cell_data_header{4}='rx_bw_mhz';
    cell_data_header{5}='rx_height';
    cell_data_header{6}='ant_hor_beamwidth';
    cell_data_header{7}='min_azimuth';
    cell_data_header{8}='max_azimuth';
    cell_data_header{9}='rx_ant_gain_mb';
    cell_data_header{10}='rx_nf';
    cell_data_header{11}='in_ratio';
    cell_data_header{12}='min_ant_loss';
    cell_data_header{13}='fdr_dB';
    cell_data_header{14}='dpa_threshold';
    cell_data_header{15}='required_pathloss';
    cell_data_header{16}='base_protection_pts';
    cell_data_header{17}='base_polygon';
    cell_data_header{18}='gmf_num';
    cell_data_header{19}='rx_lat';
    cell_data_header{20}='rx_lon';
    cell_data_header{21}='base_polyshape';
    cell_data_header{22}='ant_diamter_m';
    cell_data_header{23}='Sat_ID';
    cell_data_header{24}='Noise_TempK';
    cell_data_header{25}='Ground_Elevation_m';
    cell_data_header{26}='Antenna_Pattern_Str';
    cell_data_header{27}='rx_if_bw_mhz';
    cell_data_header{28}='array_ant_pattern';  %%%Change this to tf_custom_ant_pattern
    cell_data_header{29}='TF_Custom_Ant_Pattern';
    cell_data_header{30}='X_POL_dB';
    cell_data_header{31}='gs_azimuth';
    cell_data_header{32}='gs_elevation';
    cell_data_header{33}='tf_ant_square';     %tf_ant_square=0 %%%%%%%Instead of the trapezoid method used in CBRS
    cell_data_header{34}='second_in_ratio';
    cell_data_header{35}='second_mc_percentile';
    cell_data_header{36}='mc_percentile';
    cell_data_header{37}='dpa_second_threshold';
    cell_data_header{38}='azimuth_step';
    cell_data_header{39}='tf_keyhole_ant';  %%%%%%For part 0 to limit
    cell_data_header{40}='tf_ue_grid'; %%%%%%%If 0, then base stations
    cell_data_header{41}='ue_grid_km';  %%%%%Spacing of UEs, The sim_radius_km will set the radius
    cell_data_header{42}='ue_height_m';  %%%%%UE height in meters
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Load Loacations
    tf_pull_excel=1%0%1%0%1
    excel_filename_p2p='China Lake p2p 8-3-2026.xlsx'
    data_p2p_num=4%3%2
    mat_filename_p2p=strcat('ChinaLake_cell_p2p_adjacent_',num2str(data_p2p_num),'.mat');
    tic;
    [cell_p2p_data]=load_full_excel_rev1(app,mat_filename_p2p,excel_filename_p2p,tf_pull_excel);
    toc;
    data_header_p2p=cell_p2p_data(1,:)'

    %%%%%%%%If a number, convert to string (for the name).
    cell_p2p_data(:,1) = cellfun(@(x) num2str(x), cell_p2p_data(:,1), 'UniformOutput', false);

    for k=1:1:length(cell_p2p_data(:,1))
        %cell_p2p_data{k,1}=convertStringOrNumeric(cell_p2p_data{k,1});
    end

    sort(cell_p2p_data(:,1))
    if length(cell_p2p_data(:,1))~=length(unique(cell_p2p_data(:,1)))
        'Unique Name Error'
        pause;
    else
        'All unique names'
    end


    %%%
    %%%%%%%%%%%%%%%%'Need to pull the channel and the bandwidth'
    tf_pull_excel_chan=1%0%1%0%1
    excel_filename_chan='7ghz_p2p_channels_bandwidth.xlsx'
    data_chan_num=3%2
    mat_filename_chan=strcat('cell_p2p_channels_',num2str(data_chan_num),'.mat');
    tic;
    [cell_p2p_chan]=load_full_excel_rev1(app,mat_filename_chan,excel_filename_chan,tf_pull_excel_chan);
    toc;
    [num_chan,~]=size(cell_p2p_chan);
    for k=1:1:num_chan
        temp_band=cell_p2p_chan{k,2};

        if contains(temp_band,'M')
            temp_band=erase(temp_band,'M');
        end
        cell_p2p_chan{k,2}=str2num(temp_band);
    end
    cell_p2p_chan




    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%Stitch the data together.
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    col_data_ulink_idx=find(matches(data_header_p2p,'ULINK_ID'));
    col_data_asite_id_idx=find(matches(data_header_p2p,'A_SITE_ID'));
    col_data_asite_name_idx=find(matches(data_header_p2p,'A_SITE_NAME'));
    col_data_alat_idx=find(matches(data_header_p2p,'A_LAT_WGS84'));
    col_data_alon_idx=find(matches(data_header_p2p,'A_LON_WGS84'));
    col_data_ach1_idx=find(matches(data_header_p2p,'A_CH1'));
    col_data_afreq1_idx=find(matches(data_header_p2p,'A_FREQ1'));
    col_data_rad_ed_idx=find(matches(data_header_p2p,'BW'));
    col_data_aheight_idx=find(matches(data_header_p2p,'A_Site_Ht'));
    col_data_again_idx=find(matches(data_header_p2p,'A_Gain'));

   
    %%%'At this point, lets find the lower frequency location and put it into another cell'
    [num_p2p_rows,~]=size(cell_p2p_data);
    cell_selected_data=cell(num_p2p_rows,1);
    %%%%%%%%1) Ulink
    %%%%%%%%2) Lower frequency (LF) Site ID
    %%%%%%%%3) LF Site Name
    %%%%%%%4)  LF Lat
    %%%%%%%5) LF Lon
    %%%%%%%6) LF Channel
    %%%%%%%7) LF Frquency
    %%%%%%8) EMS Bandwidth Number MHz
    %%%%%%9) Azimuth pointing direction of receiver
    %%%%%%10) Antenna Gain
    %%%%%%11) Antenna Height
    for row_idx=2:1:num_p2p_rows
        temp_single_data=cell_p2p_data(row_idx,:);
        if isnumeric(temp_single_data{col_data_ulink_idx})
            temp_ulink=num2str(temp_single_data{col_data_ulink_idx});
        else
            temp_ulink=temp_single_data{col_data_ulink_idx};
        end

        if contains(temp_ulink,'/')
            temp_ulink=erase(temp_ulink,'/')
            temp_ulink
            'Slash Error'
            pause;
        end
        % if contains(temp_ulink,'-')
        %     temp_ulink=erase(temp_ulink,'-');
        % end
        % if contains(temp_ulink,'(')
        %     temp_ulink=erase(temp_ulink,'(');
        % end
        % if contains(temp_ulink,')')
        %     temp_ulink=erase(temp_ulink,')');
        % end
        % if contains(temp_ulink,'.')
        %     temp_ulink=erase(temp_ulink,'.');
        % end
        % if contains(temp_ulink,'_')
        %     temp_ulink=erase(temp_ulink,'_');
        % end

        %%%%%%%First check the channel, if you cant find it then use the EMS
        temp_chan=temp_single_data{col_data_ach1_idx};
        chan_row_idx=find(contains(cell_p2p_chan(:,1),temp_chan(1)));
        if length(chan_row_idx)==1
            temp_num_bw=cell_p2p_chan{chan_row_idx,2};
        elseif isempty(chan_row_idx)
            %'Legacy channel plan'
            temp_bw_ems=temp_single_data{col_data_rad_ed_idx};
            if isnumeric(temp_bw_ems)
                temp_num_bw=temp_bw_ems;
            else
                if contains(temp_bw_ems,'NaN')
                    'Nan and no channel'
                    pause;
                else
                    %'need to convert to a mhz number'
                    if contains(temp_bw_ems,'M')
                        temp_split1=strsplit(temp_bw_ems,'M');
                        temp_num_bw=str2num(temp_split1{1});
                    else
                        temp_bw_ems
                        'No M'
                        pause;
                    end
                end
            end
        else
            'Unknown branch'
            pause;
        end

    %%%%%%%%1) Ulink
    %%%%%%%%2) Lower frequency (LF) Site ID
    %%%%%%%%3) LF Site Name
    %%%%%%%4)  LF Lat
    %%%%%%%5) LF Lon
    %%%%%%%6) LF Channel
    %%%%%%%7) LF Frquency
    %%%%%%8) EMS Bandwidth Number MHz
    %%%%%%9) Azimuth pointing direction of receiver
    %%%%%%10) Antenna Gain
    %%%%%%11) Antenna Height
        temp_cell=cell(1,11);
        temp_cell{1,1}=strcat(removeSpacesAndCommas(temp_ulink));%,num2str(row_idx));
        temp_cell{1,2}=removeSpacesAndCommas(strcat(temp_single_data{col_data_asite_id_idx}));
        temp_cell{1,3}=removeSpacesAndCommas(strcat(temp_single_data{col_data_asite_name_idx}));
        temp_cell{1,4}=temp_single_data{col_data_alat_idx};
        temp_cell{1,5}=temp_single_data{col_data_alon_idx};
        temp_cell{1,6}=removeSpacesAndCommas(strcat(temp_single_data{col_data_ach1_idx}));
        temp_cell{1,7}=temp_single_data{col_data_afreq1_idx};
        temp_cell{1,8}=temp_num_bw;
        %temp_cell{1,9}=azi_A;
        temp_cell{1,10}=temp_single_data{col_data_again_idx};
        temp_cell{1,11}=temp_single_data{col_data_aheight_idx};

        cell_selected_data{row_idx,1}=temp_cell;
    end
    cell_selected_data=vertcat(cell_selected_data{:});
    cell_selected_data=cell_selected_data(~cellfun('isempty',cell_selected_data(:,1)),:);
    if length(unique(cell_selected_data(:,1)))~= length(cell_selected_data(:,1)) %%%Need to check
        'Error, not unique names'
        pause;
    end

    %%%%%%%%%%%%%%%%Example Code with OOBE and FDR
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Inputs
    rev_obbe=1;
    tx_bw_mhz=100; %%%%%%%% carrier bandwidth [MHz] [B]
    center_freq=7350; %%%%%%%MHz % carrier center [MHz]
    zero_freq=7400;
    FreqMHz=center_freq; %%%%%%%%MHz
    edge1=center_freq+tx_bw_mhz/2;
    oobeSeg = [edge1  edge1+20   -13  ;     % band edge to 7420
        edge1+20  edge1+40   -30;    % 7420-7430 MHz
        edge1+30  Inf   -40  ];    % above 7430 MHz
    cond_pwr_dBm=37.5;  % conducted power per transceiver [dBm], which is different than the 38.6dBm/1MHz or 58.6dBm/100MHz
    tx_extrap_loss=-60; %%%%%%%%%TX Extrapolation Slope dB/Decade -60dB (Past the last point 7600Mhz)
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Calculate the OOBE
    [table_oobe]=calc_oobe_7ghz_rev2_no_fig(app,rev_obbe,oobeSeg,tx_bw_mhz,center_freq,cond_pwr_dBm);
    data_header_oobe=table_oobe.Properties.VariableNames;
    cell_oobe_data=table2cell(table_oobe);
    col_freq_offset_idx=find(matches(data_header_oobe,'offset_from_center_MHz'));
    col_eirp_idx=find(matches(data_header_oobe,'EIRP_PSD_dBm_per_MHz'));
    array_mask=cell2mat(cell_oobe_data(:,[col_freq_offset_idx,col_eirp_idx]));
    %%%%%%%Normalize mask for FDR
    norm_array_mask=array_mask;
    norm_array_mask(:,2)=abs(array_mask(:,2)-max(array_mask(:,2)));
    % %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%FDR Inputs
    array_tx_rf=fliplr(norm_array_mask(:,1)'); %%%%Frequency MHz (Base Station) [Half Bandwidth]
    array_tx_mask=fliplr(norm_array_mask(:,2)'); %%%%%%%dB Loss
    tx_freq_mhz=center_freq;
    max_freq=max(cell2mat(cell_selected_data(:,7)));
    fdr_freq_separation=abs(tx_freq_mhz-max_freq);
    fdr_calc_mhz=ceil(fdr_freq_separation*1.2);
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    %%%%%%Find the unique channels and the FDR
    %%%%%%%%%%%%Just a check to make sure all the A10 channels are just 30
    uni_channels=table2cell(unique(cell2table(cell_selected_data(:,[6,7,8])),'rows'))
    %%uni_bw=unique(cell2mat(cell_selected_data(:,8)))
    uni_bw=30%unique(vertcat(cell_p2p_chan{:,2}));    %%'only calculate fdr for these'
    cut_idx=find(uni_bw<5);
    uni_bw(cut_idx)=[];
    uni_bw

    num_bw=length(uni_bw);
    cell_bw_data=cell(num_bw,4);  %%%1)MHz channel size, 2)RX IF, 3)Rx Loss 4)FDR array, freq X MHz
    norm_rx_if=fliplr(horzcat(0,15,15.1,18,30,60,90,120,150,180,210))/30; %%%%Frequency MHz Half Bandwidth
    array_rx_loss=fliplr(horzcat(0,0.1,15,30,70,110,150,190,230,270,310)); %%%%%%%dB Loss
    rx_extrap_loss=-80; %%%%%%%%%RX Extrapolation Slope dB/Decade 60dB (This is generous)
    for n=1:1:num_bw
        cell_bw_data{n,1}=uni_bw(n);
        array_rx_if=norm_rx_if*uni_bw(n);
        cell_bw_data{n,2}=array_rx_if;
        cell_bw_data{n,3}=array_rx_loss;

        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Calculate FDR
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Delta_Freq_Step=1;
        tic;
        [FDR_dB,ED,VD,OTR,DeltaFreq,single_fdr_loss,trans_mask]=FDR_ModelII_vector_rev2(app,fdr_calc_mhz,array_tx_rf,array_rx_if,array_tx_mask,array_rx_loss,tx_extrap_loss,rx_extrap_loss,Delta_Freq_Step);
        toc;

        zero_idx=nearestpoint_app(app,0,DeltaFreq);
        array_fdr=horzcat(DeltaFreq(zero_idx:end)',FDR_dB(zero_idx:end));
        cell_bw_data{n,4}=array_fdr; %%%%%%Frequency, FDR Loss

        % figure;
        % plot(array_fdr(:,1),array_fdr(:,2),'-b')
        % grid on;
        % pause(0.1)
        % 'check'
        % pause;
    end
    cell_bw_data=cell_bw_data(~cellfun('isempty',cell_bw_data(:,1)),:);


    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    col_header_label_idx=find(matches(cell_data_header,'data_label1'));
    col_header_lat_idx=find(matches(cell_data_header,'latitude'));
    col_header_lon_idx=find(matches(cell_data_header,'longitude'));
    col_header_rx_height_idx=find(matches(cell_data_header,'rx_height'));
    col_header_base_ppts_idx=find(matches(cell_data_header,'base_protection_pts'));
    col_header_base_poly_idx=find(matches(cell_data_header,'base_polygon'));
    col_header_ant_bm_idx=find(matches(cell_data_header,'ant_hor_beamwidth'));
    col_header_min_azi_idx=find(matches(cell_data_header,'min_azimuth'));
    col_header_max_azi_idx=find(matches(cell_data_header,'max_azimuth'));
    col_header_ant_gain_idx=find(matches(cell_data_header,'rx_ant_gain_mb'));
    col_header_ant_dia_idx=find(matches(cell_data_header,'ant_diamter_m'));
    col_header_rx_noise_temp_idx=find(matches(cell_data_header,'Noise_TempK'));
    col_header_rx_if_bw_idx=find(matches(cell_data_header,'rx_if_bw_mhz'));
    col_header_sat_id_idx=find(matches(cell_data_header,'Sat_ID'));
    col_header_ant_pat_str_idx=find(matches(cell_data_header,'Antenna_Pattern_Str'));
    col_header_tf_cust_ant_idx=find(matches(cell_data_header,'TF_Custom_Ant_Pattern'));
    col_header_x_pol_dB_idx=find(matches(cell_data_header,'X_POL_dB'));
    col_header_in_ratio_idx=find(matches(cell_data_header,'in_ratio'));
    col_header_fdr_dB_idx=find(matches(cell_data_header,'fdr_dB'));
    col_header_tf_ant_idx=find(matches(cell_data_header,'tf_ant_square'));
    col_header_second_IN_idx=find(matches(cell_data_header,'second_in_ratio'));
    col_header_second_per_idx=find(matches(cell_data_header,'second_mc_percentile'));
    col_header_mc_per_idx=find(matches(cell_data_header,'mc_percentile'));
    col_header_key_idx=find(matches(cell_data_header,'tf_keyhole_ant'));
    col_header_step_azimuth_idx=find(matches(cell_data_header,'azimuth_step'));
    col_header_gs_azimuth_idx=find(matches(cell_data_header,'gs_azimuth'));
    col_header_gs_elevation_idx=find(matches(cell_data_header,'gs_elevation'));
    col_header_min_ant_idx=find(matches(cell_data_header,'min_ant_loss'));
    
    % 'Just put all points into 1'
    % pause;

    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    [num_locations,~]=size(cell_selected_data);
    num_locations=1
    %table([1:num_locations]',cell_selected_data(:,1))
    tic;
    cell_compiled_data=cell(num_locations,42);
    cell_compiled_data(:,col_header_min_azi_idx)=num2cell(0); %%%Double Zero should keep it fixed (Non-rotating)
    cell_compiled_data(:,col_header_max_azi_idx)=num2cell(0); %%%Double Zero should keep it fixed (Non-rotating)
    cell_compiled_data(:,col_header_ant_bm_idx)=num2cell(360); %%%%%%%%omni directional

    cell_compiled_data(:,col_header_min_ant_idx)=num2cell(0);
    
    cell_compiled_data(:,col_header_tf_cust_ant_idx)=num2cell(0);
    cell_compiled_data(:,col_header_x_pol_dB_idx)=num2cell(2); %%%%x_pol_dB 2dB
    cell_compiled_data(:,col_header_in_ratio_idx)=num2cell(-10); %%%I/N dB
    cell_compiled_data(:,col_header_key_idx)=num2cell(0);
    cell_compiled_data(:,col_header_step_azimuth_idx)=num2cell(0);
    cell_compiled_data(:,col_header_gs_elevation_idx)=num2cell(0);
    cell_compiled_data(:,col_header_tf_ant_idx)=num2cell(0);

    array_bw=cell2mat(cell_bw_data(:,1));
    for base_idx=1%:1:num_locations
        temp_single_cell_sim_data=cell_selected_data(base_idx,:);
        data_label1=temp_single_cell_sim_data{1};
        if contains(data_label1,'/')
            data_label1=erase(data_label1,'/')
            data_label1
            'Slash Error'
            pause;
        end

        %%%%%%%%Insert Agency Name
        %data_label1=data_label1(find(~isspace(data_label1)));  %%%%%%%%%%Remove the White Spaces
        cell_compiled_data{base_idx,col_header_label_idx}=strcat(data_label1);%,num2str(base_idx));

        %%%%%%%%%%Lat/Lon/Height
                %%%%All points
        % % 'all points'
        % % pause;

        
        temp_lat=vertcat(cell_selected_data{:,4}) %%%convertStringOrNumeric(temp_single_cell_sim_data{4});
        temp_lon=vertcat(cell_selected_data{:,5}) %%convertStringOrNumeric(temp_single_cell_sim_data{5});
        cell_compiled_data{base_idx,col_header_lat_idx}=temp_lat;
        cell_compiled_data{base_idx,col_header_lon_idx}=temp_lon;
        %%temp_ant_height=30;
        temp_ant_height=temp_single_cell_sim_data{11};
        temp_ant_gain=temp_single_cell_sim_data{10};
        cell_compiled_data{base_idx,col_header_ant_gain_idx}=temp_ant_gain;
        
        %%%%%%10) Antenna Gain
        %%%%%%11) Antenna Height
        
        % if isnan(temp_ant_height)
        %     temp_ant_height=10
        % end

        cell_compiled_data{base_idx,col_header_rx_height_idx}=temp_ant_height.*ones(size(temp_lat));
        cell_compiled_data{base_idx,col_header_base_ppts_idx}=horzcat(cell_compiled_data{:,col_header_lat_idx},cell_compiled_data{base_idx,col_header_lon_idx},cell_compiled_data{base_idx,col_header_rx_height_idx});
        cell_compiled_data{base_idx,col_header_base_poly_idx}=horzcat(cell_compiled_data{base_idx,col_header_lat_idx},cell_compiled_data{base_idx,col_header_lon_idx});

        %%%%%%%%%%FDR
        temp_freq_sep=abs(center_freq-temp_single_cell_sim_data{7});
        temp_bw=temp_single_cell_sim_data{8};
        %%%%%%%%%%%Receiver IF Bandwidth MHz
        cell_compiled_data{base_idx,col_header_rx_if_bw_idx}=temp_bw;

        %%%%Find the closest
        nn_data_row_idx=nearestpoint_app(app,temp_bw,array_bw);
        single_bw_data=cell_bw_data(nn_data_row_idx,:);
        temp_fdr_array=single_bw_data{4};
        nn_fdr_row_idx=nearestpoint_app(app,temp_freq_sep,temp_fdr_array(:,1));
        cell_compiled_data(base_idx,col_header_fdr_dB_idx)=num2cell(temp_fdr_array(nn_fdr_row_idx,2)); %%%FDR dB

        %%%%%%%%%%%Azimuth Pointing
        cell_compiled_data(base_idx,col_header_gs_azimuth_idx)=num2cell(0);
    end
    cell_compiled_data=cell_compiled_data(~cellfun('isempty',cell_compiled_data(:,1)),:);
    cell_compiled_data=cell_compiled_data(~cellfun('isempty',cell_compiled_data(:,2)),:);

    num_locations
    size(cell_compiled_data)
    cell_sim_data=vertcat(cell_data_header,cell_compiled_data);

    cell_sim_data(1:2,:)'
    tic;
    save(cell_data_filename,'cell_sim_data')
    toc;
end
cell_sim_data(1:2,:)'





'Now lets try to do an example'
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Test Simulation
rev=116; %%%%%%China Lake Test (8 pts: Multi-pts) 
size(cell_sim_data) 
%%%%%%%%%bs_eirp=51.75%%%% %%%%%EIRP [dBm/50MHz]  %%%(old value with 50mhz)
tx_eirp_mhz=42.5  %%%%42.5dBm/MHz at 100th (0,0)  71%65  %%%%%dBm/1Mhz
tx_bw_mhz=100; %%%%%%%% carrier bandwidth [MHz] [B]
bs_eirp=tx_eirp_mhz+10*log10(tx_bw_mhz/1) %%%%%%%90dBm/100Mhz
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
grid_spacing=0.1%0.25%1%2%3%10%5%10;  %%%%km:
sim_scale_factor=5%4%3%2; %%%%%%%%%%The scaling Factor for the simulation area. 1.1 is a 10% increase in distance post ITM
max_itm_dist_km=500; %%%%%It just makes it easy if we have a max number
mitigation_dB=0%:10:30;  %%%%%%%%% in dB%%%%% Beam Muting or PRB Blanking (or any other mitigation mechanism):  30 dB reduction %%%%%%%%%%%%Consider have this be an array, 3dB step size, to get a more granular insight into how each 3dB mitigation reduces the coordination zone.
reliability=50%20%10%50%
Tpol=1; %%%polarization for ITM
FreqMHz=7125;
confidence=50;
tx_height_m=25%NaN(1,1); %%%%If NaN, then keep the normal heights: base_station_height
tf_clutter=0;%1;  %%%%%%This if for P2108.
sim_folder1='Z:\Matlab2025 Sims\7GHz P2P Bugsplats'  %%%%%'C:\Local Matlab Data\7GHz P2P Single Entry'%%%
%%%%%%%%%%%%%%%%%%%%%%%P2P
in_ratio=-10; %%%%%I/N Ratio
rx_temp_k=293;%%%%%Noise Temperature K
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

cell_sim_data'




%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%Create a Rev Folder
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
cd(sim_folder1);
pause(0.1)
tempfolder=strcat('Rev',num2str(rev));
[status,msg,msgID]=mkdir(tempfolder);
rev_folder=fullfile(sim_folder1,tempfolder);
cd(rev_folder)
pause(0.1)


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%Base Station EIRP Data
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
norm_aas_zero_elevation_data=zeros(361,4);
norm_aas_zero_elevation_data(:,1)=-180:1:180;
%%%%1) Azimuth -180~~180
%%%2) Rural
%%%3) Suburban
%%%4) Urban
%%%%AAS Reduction in Gain to Max Gain (0dB is 0dB reduction, which equates to the make antenna gain of 25dB)
%%%%Need to normalize to zero after the "downtilt reductions" are calculated
%%%%To simplify the data, this is gain at the horizon. 50th Percentile
'Might need to pull the EIRP mask if we want to do a different percentile'
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
bs_down_tilt_reduction=0;
bs_eirp_reductions=(bs_eirp-bs_down_tilt_reduction)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Saving the simulation files in a folder for the option to run from a server
'First save . . .' %%%%%24 seconds on Z drive
tic;
save('grid_spacing.mat','grid_spacing')
save('reliability.mat','reliability')
save('confidence.mat','confidence')
save('FreqMHz.mat','FreqMHz')
save('Tpol.mat','Tpol')
save('norm_aas_zero_elevation_data.mat','norm_aas_zero_elevation_data')
save('tf_clutter.mat','tf_clutter')
save('mitigation_dB.mat','mitigation_dB')
save('tx_height_m.mat','tx_height_m')
save('bs_eirp_reductions.mat','bs_eirp_reductions')
save('sim_scale_factor.mat','sim_scale_factor')
toc;


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Should probably pull this out of the loop so we don't have to do it 4,000 times
%%%%%%%%%%Find the ITM Area Pathloss for the distance array
tic;
max_rx_height=30%max(cell2mat(cell_sim_data(:,5)))
[array_dist_pl]=itm_area_dist_array_sea_rev2(app,reliability,tx_height_m,max_rx_height,max_itm_dist_km,FreqMHz);
toc;
tic;
save(strcat('Rev',num2str(rev),'_array_dist_pl.mat'),'array_dist_pl')
toc;
%%%'save the array_dist_pl as the server will use this to create the azi array'

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%First loop does all the calculation for the 15 columns, then just saves the cell_sim_data for the server to make the folders
%%%%%%%%%For Loop the Locations

cell_data_header=cell_sim_data(1,:);
col_data_label_idx=find(matches(cell_data_header,'data_label1'));
col_dpa_threshold_idx=find(matches(cell_data_header,'dpa_threshold'));
col_req_path_idx=find(matches(cell_data_header,'required_pathloss'));
col_header_fdr_dB_idx=find(matches(cell_data_header,'fdr_dB'));
col_header_rx_if_bw_idx=find(matches(cell_data_header,'rx_if_bw_mhz'));
col_header_x_pol_dB_idx=find(matches(cell_data_header,'X_POL_dB'));
col_header_ant_gain_idx=find(matches(cell_data_header,'rx_ant_gain_mb'));
cell_sim_data'

[num_locations,~]=size(cell_sim_data);
table([1:num_locations]',cell_sim_data(:,1))
tic;
for base_idx=2:1:num_locations
    temp_single_cell_sim_data=cell_sim_data(base_idx,:);
    temp_fdr_dB=temp_single_cell_sim_data{col_header_fdr_dB_idx}
    rx_bw_mhz=temp_single_cell_sim_data{col_header_rx_if_bw_idx};
    x_pol_dB=temp_single_cell_sim_data{col_header_x_pol_dB_idx};
    rx_ant_gain=temp_single_cell_sim_data{col_header_ant_gain_idx}
    dpa_threshold=-138.7+10*log10(rx_bw_mhz)+10*log10(rx_temp_k)+in_ratio+x_pol_dB+-rx_ant_gain+temp_fdr_dB;  %%%%%%We normalize the antenna pattern
    cell_sim_data{base_idx,col_dpa_threshold_idx}=dpa_threshold;
    required_pathloss=ceil(bs_eirp_reductions-dpa_threshold)
    cell_sim_data{base_idx,col_req_path_idx}=required_pathloss;
end
toc;  %%%%%%%%%%%Very quick.


cd(rev_folder)
pause(0.1)
cell_sim_data
cell_sim_data(1:2,:)'
'Last save . . .'
tic;
save('cell_sim_data.mat','cell_sim_data')
toc;






%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Now running the simulation
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
tf_server_status=0;
parallel_flag=0%1%0;
[workers,parallel_flag]=check_parallel_toolbox(app,parallel_flag)
workers=2
tf_recalculate=0
tf_rescrap_rev_data=1
tf_print_excel=1%0%1




%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
wrapper_bugsplat_merge_folders_geoplot_pea2023_p2p_rev17(app,rev_folder,parallel_flag,workers,tf_server_status,tf_recalculate,tf_rescrap_rev_data)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%





%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



end_clock=clock;
total_clock=end_clock-top_start_clock;
total_seconds=total_clock(6)+total_clock(5)*60+total_clock(4)*3600+total_clock(3)*86400;
total_mins=total_seconds/60;
total_hours=total_mins/60;
if total_hours>1
    strcat('Total Hours:',num2str(total_hours))
elseif total_mins>1
    strcat('Total Minutes:',num2str(total_mins))
else
    strcat('Total Seconds:',num2str(total_seconds))
end
cd(folder1)
'Done'

rev_folder









