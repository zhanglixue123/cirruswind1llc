%% Analysis Report

% 求总共盈亏各多少
excel = 'F:\Cirrus Wind 1\xcelenergy\2018-06 Cirrus Support.xlsx ';   %2017-05 Cirrus support可被替换
revenue=xlsread(excel,'2018-06 Cirrus Support','k2:k8641 '); %'k2:k8929'根据当月的情况重新填
revenue_positive = revenue(find (revenue>=0));
revenue_negative = revenue(find (revenue<0));
sum_total = sum(revenue);
sum_positive = sum(revenue_positive);
sum_negative = sum(revenue_negative);

% 赋值
SSave = 0;
LLoss = 0;

EstimateKW0 = 0;
% 把EstimateKW0复制粘贴到变量里
LMP=0;
% 把LMP复制粘贴到变量里
EnergyCost = 0;
% 把实际的EnergyCost复制粘贴到变量里

% 算节省的钱
KW0=mean(EstimateKW0(:));
LMP_negative = LMP(find (LMP<0));   
EnergyCost0 = LMP_negative * KW0;   % 0代表假的
EnergyCost1 = EnergyCost(find (EnergyCost<0)) ; % 1代表真的
sumEnergyCost0 = sum(EnergyCost0);
sumEnergyCost1 = sum(EnergyCost1);
Save = sumEnergyCost1 - sumEnergyCost0;
% 把Save依次复制到SSave变量中

% 算赔了的钱
KW0=mean(EstimateKW0(:));
LMP_positive = LMP(find (LMP>=0));
EnergyCost00 = LMP_positive * KW0;
EnergyCost11 = EnergyCost(find (EnergyCost>=0)) ;
sumEnergyCost00 = sum(EnergyCost00);
sumEnergyCost11 = sum(EnergyCost11);
Loss = sumEnergyCost00 - sumEnergyCost11;
% 把Loss依次复制到LLoss变量中

Total_Save = sum(SSave); %算节省的钱之和
Total_Loss = sum(LLoss); %算赔了的钱之和




%% 出力太小的时间段
excel = 'F:\Cirrus Wind 1\xcelenergy\2018-06 Cirrus Support.xlsx';   %2017-05 Cirrus support可被替换
output = xlsread(excel,'2018-06 Cirrus Support','g2:g8641'); %'g2:g8929'根据当月的情况重新填，注意是g不是k了
output_low = output(find (output<50));  % 找当月出力少于50的
output_low_hr = length(output_low)*5/60;




%% 算算我该几点睡

% 初始值
sum_actualincome  = 0;  %   把算好的sum_actualincome复制到变量里

LLoss = 0;

% 1. 如果出力完全60000

EstimateKW0 = 0;
% 把EstimateKW0复制粘贴到变量里
LMP=0;
% 把LMP复制粘贴到变量里
EnergyCost = 0;
% 把实际的EnergyCost复制粘贴到变量里

KW0 = mean(EstimateKW0(:));
EnergyCost00 = LMP * KW0; %把2000的部分变成60000
EnergyCost11 = EnergyCost; % 真·收入
sumEnergyCost00 = sum(EnergyCost00);
sumEnergyCost11 = sum(EnergyCost11);
Loss = sumEnergyCost00 - sumEnergyCost11;
%把Loss粘贴到LLoss中

sum_LLoss = sum(LLoss);
income6w = sum_actualincome + sum_LLoss; % 假设全部为60000的总收入

% 2. 如果出力完全2000

% 初始值
sum_actualincome  = 0;  %   把算好的sum_actualincome复制到变量里

SSave = 0;

EstimateKW0 = 0;
% 把EstimateKW0复制粘贴到变量里

LMP=0;
% 把LMP复制粘贴到变量里
EnergyCost = 0;
% 把实际的EnergyCost复制粘贴到变量里

KW0 = mean(EstimateKW0(:));

EnergyCost0 = LMP * KW0; %把60000的部分变成2000
EnergyCost1 = EnergyCost; % 真·收入
sumEnergyCost0 = sum(EnergyCost0);
sumEnergyCost1 = sum(EnergyCost1);
Save = sumEnergyCost1 - sumEnergyCost0;
%把Save粘贴到SSave中

%批量计算
sum_SSave = Save' ;
income2k = sum_actualincome - sum_SSave; % 假设全部为2000的总收入

%不批量计算
sum_SSave = sum(SSave);
income2k = sum_actualincome - sum_SSave; % 假设全部为2000的总收入

% 如果没有调整风机，KW0计算
excel = 'F:\Cirrus Wind 1\xcelenergy\2018-06 Cirrus Support.xlsx';   %2017-05 Cirrus support可被替换
output = xlsread(excel,'2018-06 Cirrus Support','g2:g8641'); %'g2:g8929'根据当月的情况重新填，注意是g不是k了
KW0_up = output(find (output<100));  
KW0_down = KW0_up(find (KW0_up>45));  
KW0 = mean(KW0_down(:));












