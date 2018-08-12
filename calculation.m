%% Analysis Report

% ���ܹ�ӯ��������
excel = 'F:\Cirrus Wind 1\xcelenergy\2018-06 Cirrus Support.xlsx ';   %2017-05 Cirrus support�ɱ��滻
revenue=xlsread(excel,'2018-06 Cirrus Support','k2:k8641 '); %'k2:k8929'���ݵ��µ����������
revenue_positive = revenue(find (revenue>=0));
revenue_negative = revenue(find (revenue<0));
sum_total = sum(revenue);
sum_positive = sum(revenue_positive);
sum_negative = sum(revenue_negative);

% ��ֵ
SSave = 0;
LLoss = 0;

EstimateKW0 = 0;
% ��EstimateKW0����ճ����������
LMP=0;
% ��LMP����ճ����������
EnergyCost = 0;
% ��ʵ�ʵ�EnergyCost����ճ����������

% ���ʡ��Ǯ
KW0=mean(EstimateKW0(:));
LMP_negative = LMP(find (LMP<0));   
EnergyCost0 = LMP_negative * KW0;   % 0����ٵ�
EnergyCost1 = EnergyCost(find (EnergyCost<0)) ; % 1�������
sumEnergyCost0 = sum(EnergyCost0);
sumEnergyCost1 = sum(EnergyCost1);
Save = sumEnergyCost1 - sumEnergyCost0;
% ��Save���θ��Ƶ�SSave������

% �����˵�Ǯ
KW0=mean(EstimateKW0(:));
LMP_positive = LMP(find (LMP>=0));
EnergyCost00 = LMP_positive * KW0;
EnergyCost11 = EnergyCost(find (EnergyCost>=0)) ;
sumEnergyCost00 = sum(EnergyCost00);
sumEnergyCost11 = sum(EnergyCost11);
Loss = sumEnergyCost00 - sumEnergyCost11;
% ��Loss���θ��Ƶ�LLoss������

Total_Save = sum(SSave); %���ʡ��Ǯ֮��
Total_Loss = sum(LLoss); %�����˵�Ǯ֮��




%% ����̫С��ʱ���
excel = 'F:\Cirrus Wind 1\xcelenergy\2018-06 Cirrus Support.xlsx';   %2017-05 Cirrus support�ɱ��滻
output = xlsread(excel,'2018-06 Cirrus Support','g2:g8641'); %'g2:g8929'���ݵ��µ���������ע����g����k��
output_low = output(find (output<50));  % �ҵ��³�������50��
output_low_hr = length(output_low)*5/60;




%% �����Ҹü���˯

% ��ʼֵ
sum_actualincome  = 0;  %   ����õ�sum_actualincome���Ƶ�������

LLoss = 0;

% 1. ���������ȫ60000

EstimateKW0 = 0;
% ��EstimateKW0����ճ����������
LMP=0;
% ��LMP����ճ����������
EnergyCost = 0;
% ��ʵ�ʵ�EnergyCost����ճ����������

KW0 = mean(EstimateKW0(:));
EnergyCost00 = LMP * KW0; %��2000�Ĳ��ֱ��60000
EnergyCost11 = EnergyCost; % �桤����
sumEnergyCost00 = sum(EnergyCost00);
sumEnergyCost11 = sum(EnergyCost11);
Loss = sumEnergyCost00 - sumEnergyCost11;
%��Lossճ����LLoss��

sum_LLoss = sum(LLoss);
income6w = sum_actualincome + sum_LLoss; % ����ȫ��Ϊ60000��������

% 2. ���������ȫ2000

% ��ʼֵ
sum_actualincome  = 0;  %   ����õ�sum_actualincome���Ƶ�������

SSave = 0;

EstimateKW0 = 0;
% ��EstimateKW0����ճ����������

LMP=0;
% ��LMP����ճ����������
EnergyCost = 0;
% ��ʵ�ʵ�EnergyCost����ճ����������

KW0 = mean(EstimateKW0(:));

EnergyCost0 = LMP * KW0; %��60000�Ĳ��ֱ��2000
EnergyCost1 = EnergyCost; % �桤����
sumEnergyCost0 = sum(EnergyCost0);
sumEnergyCost1 = sum(EnergyCost1);
Save = sumEnergyCost1 - sumEnergyCost0;
%��Saveճ����SSave��

%��������
sum_SSave = Save' ;
income2k = sum_actualincome - sum_SSave; % ����ȫ��Ϊ2000��������

%����������
sum_SSave = sum(SSave);
income2k = sum_actualincome - sum_SSave; % ����ȫ��Ϊ2000��������

% ���û�е��������KW0����
excel = 'F:\Cirrus Wind 1\xcelenergy\2018-06 Cirrus Support.xlsx';   %2017-05 Cirrus support�ɱ��滻
output = xlsread(excel,'2018-06 Cirrus Support','g2:g8641'); %'g2:g8929'���ݵ��µ���������ע����g����k��
KW0_up = output(find (output<100));  
KW0_down = KW0_up(find (KW0_up>45));  
KW0 = mean(KW0_down(:));












