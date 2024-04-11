clear;

% エクセルファイルのパスを指定
[filename, filepath] = uigetfile('*.xlsx;*.xls');

% シート名を指定
sheet_name = 'Sheet2';

% エクセルファイルから数値データを読み取り
data = readmatrix(filename, 'Sheet', sheet_name, 'Range', 'C3:D30');

% 横軸と縦軸のデータを取得
x_data = data(:, 1); % セルC3〜C30を横軸とする
y_data = data(:, 2); % セルD3〜D30を縦軸とする

% 最小値と最大値を計算
min_y = min(y_data);
max_y = max(y_data);

% 数値を0〜1の範囲に正規化
normalized_y_data = (y_data - min_y) / (max_y - min_y);

% データのフィッティング
fit_result = fit(x_data, normalized_y_data, 'fourier8');

%----------
a0 = fit_result.a0;
a1 = fit_result.a1;
b1 = fit_result.b1;
a2 = fit_result.a2;
b2 = fit_result.b2;
a3 = fit_result.a3;
b3 = fit_result.b3;
a4 = fit_result.a4;
b4 = fit_result.b4;
a5 = fit_result.a5;
b5 = fit_result.b5;
a6 = fit_result.a6;
b6 = fit_result.b6;
a7 = fit_result.a7;
b7 = fit_result.b7;
a8 = fit_result.a8;
b8 = fit_result.b8;
w =  fit_result.w;
%----------
syms g(x)
g(x) = a0 + a1*cos(x*w) + b1*sin(x*w) + ...
               a2*cos(2*x*w) + b2*sin(2*x*w) + a3*cos(3*x*w) + b3*sin(3*x*w) + ...
               a4*cos(4*x*w) + b4*sin(4*x*w) + a5*cos(5*x*w) + b5*sin(5*x*w) + ...
               a6*cos(6*x*w) + b6*sin(6*x*w) + a7*cos(7*x*w) + b7*sin(7*x*w) + ...
               a8*cos(8*x*w) + b8*sin(8*x*w);
sol1 = vpasolve(g == 0.2, x, 1);
sol2 = vpasolve(g == 0.2, x, 6);
sol3 = vpasolve(g == 0.8, x, 1.5);
sol4 = vpasolve(g == 0.8, x, 6);

% ラインプロット
figure(1);
plot(fit_result, x_data, normalized_y_data);
xlabel('横軸ラベル');
ylabel('縦軸ラベル');
ylim([0, max(normalized_y_data)]);
    
% 赤いグリッド線を引く
hold on;
plot(xlim, [0.2, 0.2], 'r-'); 
plot(xlim, [0.8, 0.8], 'r-'); 

% sol1 と sol2 をプロット
plot(double(sol1), g(sol1), 'ro');
plot(double(sol2), g(sol2), 'ro');
plot(double(sol3), g(sol3), 'ro');
plot(double(sol4), g(sol4), 'ro');

% sol1 と sol2 から x 軸に垂直な線を引く
line([double(sol1), double(sol1)], [0, g(sol1)], 'Color', 'b', 'LineStyle', '--'); 
line([double(sol2), double(sol2)], [0, g(sol2)], 'Color', 'b', 'LineStyle', '--');
line([double(sol3), double(sol3)], [0, g(sol3)], 'Color', 'b', 'LineStyle', '--');
line([double(sol4), double(sol4)], [0, g(sol4)], 'Color', 'b', 'LineStyle', '--');

hold off;

sharpness = 2/ ((sol3 - sol1) + (sol2 - sol4));
title(strcat('sharpness = ', num2str(double(sharpness)), 'mm^-^1'));
